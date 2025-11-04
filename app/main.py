import json
import os
import sys
import tempfile
import traceback
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

# 将项目根目录加入 Python 搜索路径
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from service.converter import DocumentConverter

app = FastAPI(
    title="文档转换服务",
    description="将 DOCX 文档转换为标准格式的 Web 服务",
    version="1.0.0",
)

BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATE_DIR = BASE_DIR / "template"
TEMPLATES_DIR = BASE_DIR / "templates"
RESULT_DIR = BASE_DIR / "result"
TEMP_DIR = BASE_DIR / "temp"

# 挂载静态文件目录，提供前端资源
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# 配置模板目录
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

# 创建必要的目录
RESULT_DIR.mkdir(exist_ok=True)
TEMP_DIR.mkdir(exist_ok=True)


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/api")
async def root():
    return {"message": "文档转换服务正在运行", "version": "1.0.0"}


@app.get("/result/{filename}")
async def get_result_file(filename: str):
    file_path = RESULT_DIR / filename
    if file_path.exists():
        return FileResponse(
            path=str(file_path),
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    raise HTTPException(status_code=404, detail="文件未找到")


class StreamConverter:
    def __init__(self) -> None:
        self.logs = []

    def write(self, message: str) -> int:
        self.logs.append(message)
        return len(message)

    def flush(self) -> None:
        pass

    def get_logs(self) -> list[str]:
        logs = self.logs.copy()
        self.logs.clear()
        return logs


class DocumentConverterService:
    def __init__(
        self,
        file: UploadFile,
        template_file: UploadFile | None,
        header_text: str,
        toc_title: str,
        save_intermediate: bool,
        document_type: int,
    ) -> None:
        self.file = file
        self.template_file = template_file
        self.header_text = header_text
        self.toc_title = toc_title
        self.save_intermediate = save_intermediate
        self.document_type = document_type
        self.source_file_path: str | None = None
        self.template_file_path: str | None = None

    async def prepare_files(self) -> None:
        # 保存上传的源文件到临时路径
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_file:
            content = await self.file.read()
            temp_file.write(content)
            self.source_file_path = temp_file.name

        # 如果提供了模板文件，保存至临时路径
        if self.template_file and self.template_file.filename:
            if not self.template_file.filename.endswith(".docx"):
                raise HTTPException(status_code=400, detail="模板文件必须是 DOCX 格式")

            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_template_file:
                content = await self.template_file.read()
                temp_template_file.write(content)
                self.template_file_path = temp_template_file.name

    async def convert_document_stream(self):
        try:
            output_filename = f"converted_{self.file.filename}"
            output_file_path = RESULT_DIR / output_filename

            template_path = self.template_file_path
            if not template_path:
                template_file = TEMPLATE_DIR / "reference_content.docx"
                if not template_file.exists():
                    for template_name in ["reference.docx", "template.docx"]:
                        candidate = TEMPLATE_DIR / template_name
                        if candidate.exists():
                            template_file = candidate
                            break
                template_path = str(template_file) if template_file.exists() else None

            original_stdout = sys.stdout
            stream_converter = StreamConverter()
            sys.stdout = stream_converter

            try:
                with DocumentConverter(document_type=self.document_type) as converter:
                    success = converter.convert_document(
                        source_file=self.source_file_path,
                        output_file=str(output_file_path),
                        template_file=template_path,
                        header_text=self.header_text,
                        toc_title=self.toc_title,
                        save_intermediate=self.save_intermediate,
                        document_type=self.document_type,
                    )

                for log in stream_converter.get_logs():
                    yield f"data: {json.dumps({'type': 'log', 'message': log})}\n\n"

                if success and output_file_path.exists():
                    yield f"data: {json.dumps({'type': 'complete', 'filename': output_filename})}\n\n"
                else:
                    yield f"data: {json.dumps({'type': 'error', 'message': '文档转换失败'})}\n\n"

            except Exception as exc:  # noqa: BLE001
                error_msg = f"转换过程中发生错误: {exc}\n{traceback.format_exc()}"
                yield f"data: {json.dumps({'type': 'error', 'message': error_msg})}\n\n"
            finally:
                sys.stdout = original_stdout

        except Exception as exc:  # noqa: BLE001
            yield f"data: {json.dumps({'type': 'error', 'message': f'处理过程中发生错误: {exc}'})}\n\n"
        finally:
            if self.source_file_path and os.path.exists(self.source_file_path):
                os.unlink(self.source_file_path)
            if self.template_file_path and os.path.exists(self.template_file_path):
                os.unlink(self.template_file_path)


@app.post("/convert-document/")
async def convert_document(
    file: UploadFile = File(...),
    template_file: UploadFile | None = File(None),
    header_text: str = Form("格式化文档"),
    toc_title: str = Form("目录"),
    save_intermediate: bool = Form(False),
    document_type: int = Form(1),
):
    """将上传的文档转换为标准格式。"""

    converter_service = DocumentConverterService(
        file=file,
        template_file=template_file,
        header_text=header_text,
        toc_title=toc_title,
        save_intermediate=save_intermediate,
        document_type=document_type,
    )

    await converter_service.prepare_files()

    return StreamingResponse(
        converter_service.convert_document_stream(),
        media_type="text/event-stream",
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8002)
