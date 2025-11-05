import win32com.client as win32
import os
import time
from pathlib import Path

try:
    # 导入Pillow库中的ImageGrab，用于从剪贴板抓取图像
    from PIL import ImageGrab
except ImportError:
    print("错误：未找到 Pillow (PIL) 库。")
    print("请先安装: pip install Pillow")
    exit()

# --- pywin32 COM 常量 ---
wdInlineShapeEmbeddedOLEObject = 1  # 内嵌 OLE 对象

# --- 默认图片目录（项目 temp 子目录） ---
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_IMAGE_DIR = PROJECT_ROOT / "temp" / "formula_images"


def replace_ole_with_images(docx_path, output_docx_path, output_img_dir=None):
    """
    打开一个Word文档，遍历所有嵌入式 OLE 对象 (InlineShape OLE)，
    将其截图保存为 PNG，然后用该 PNG 替换掉原 OLE 对象。
    最后将修改后的文档另存为新文件。

    (第8版: 实现 OLE 对象的 "截图-删除-插入" 替换)
    """

    # 确保所有路径都是绝对路径
    if not os.path.isabs(docx_path):
        docx_path = os.path.abspath(docx_path)
    if not os.path.isabs(output_docx_path):
        output_docx_path = os.path.abspath(output_docx_path)

    if output_img_dir:
        output_img_dir = Path(output_img_dir).expanduser().resolve()
    else:
        output_img_dir = DEFAULT_IMAGE_DIR

    output_img_dir = output_img_dir.resolve()

    # 1. 确保图片输出目录存在
    output_img_dir.mkdir(parents=True, exist_ok=True)
    print(f"图片将保存到: {output_img_dir}")

    word = None
    doc = None
    saved_image_paths = []

    try:
        print("正在启动 Word 应用程序...")
        word = win32.Dispatch("Word.Application")
        word.Visible = True
        word.DisplayAlerts = False
        word.Activate()

        print(f"正在打开文档: {docx_path}")
        doc = word.Documents.Open(docx_path)
        print("文档打开成功。")

        # --- 遍历 InlineShapes 并替换 ---
        print("正在遍历 InlineShapes 查找 OLE 对象并替换...")
        shapes = doc.InlineShapes
        shape_count = shapes.Count
        found_ole_count = 0

        if shape_count > 0:
            print(f"找到 {shape_count} 个 InlineShapes，正在检查...")
            # 必须从后向前遍历，因为我们会删除和添加元素
            for i in range(shape_count, 0, -1):
                shape = None
                try:
                    shape = shapes(i)

                    # 检查是否是 OLE 对象
                    if shape.Type == wdInlineShapeEmbeddedOLEObject:

                        prog_id = "Unknown"
                        try:
                            prog_id = shape.OLEFormat.ProgID
                        except Exception:
                            pass

                        print(f"  > 找到 OLE 对象 (Shape {i}), ProgID: {prog_id}。正在替换...")

                        # 0. 【关键】获取 OLE 对象的准确位置 (Range)
                        original_range = shape.Range

                        # 1. 选中形状并复制为图片
                        shape.Select()
                        word.Selection.CopyAsPicture()
                        time.sleep(0.1)

                        # 2. 从剪贴板抓取图像
                        image = ImageGrab.grabclipboard()

                        if image:
                            # 3. 保存图像到文件
                            file_name = f"formula_shape_{i}.png"
                            save_path = output_img_dir / file_name
                            image.save(save_path)
                            saved_image_paths.append(str(save_path))

                            # 4. 【关键】删除原始 OLE 对象
                            shape.Delete()

                            # 5. 【关键】在原位置插入图片
                            # LinkToFile=False (不链接文件), SaveWithDocument=True (嵌入文档)
                            new_shape = doc.InlineShapes.AddPicture(
                                FileName=str(save_path),
                                LinkToFile=False,
                                SaveWithDocument=True,
                                Range=original_range
                            )
                            # (可选) 保持图片宽高比
                            new_shape.LockAspectRatio = True

                            found_ole_count += 1
                        else:
                            print(f"  > 警告: 复制 Shape {i} 到剪贴板失败。")

                except Exception as e_shape:
                    print(f"  > 错误: 处理 Shape {i} 时出错: {e_shape}")
                    pass

            print(f"步骤完成: 检查了 {shape_count} 个形状，成功替换 {found_ole_count} 个 OLE 对象。")

            # 6. 【关键】另存为新文档
            print(f"正在将修改后的文档另存为: {output_docx_path}")
            doc.SaveAs(output_docx_path)

        else:
            print("未找到 InlineShapes。")


    except Exception as e:
        print(f"处理文档时出错: {e}")

    finally:
        if doc:
            doc.Close()  # 关闭 (已保存的) 新文档
            print("已关闭文档。")
        if word:
            word.Quit(0)
            print("已退出 Word 应用程序。")

    return saved_image_paths


# --- 使用示例 ---
if __name__ == "__main__":

    # 1. 定义原始文件路径
    file_path = r"E:\1文档\鲅鱼圈炼焦节能泵改造可研汇1111111.docx"

    # 2. 定义新文件的保存路径 (!!!)
    # 我们可以简单地在原文件名后添加 "_replaced"
    base_name = os.path.basename(file_path)
    dir_name = os.path.dirname(file_path)
    name, ext = os.path.splitext(base_name)

    new_file_name = f"{name}_replaced{ext}"
    new_file_path = os.path.join(dir_name, new_file_name)

    if os.path.exists(file_path):
        images = replace_ole_with_images(file_path, new_file_path)

        if images:
            print("\n--- 替换完成 ---")
            image_dir = Path(images[0]).parent
            print(f"已创建 {len(images)} 张图片，保存于: {image_dir}")
            print("已将原文档中的公式替换为图片，并另存为：")
            print(new_file_path)
        else:
            print("\n(v8) 未找到可替换的 OLE 对象。")
    else:
        print(f"错误：文件未找到! \n请检查路径: {file_path}")
