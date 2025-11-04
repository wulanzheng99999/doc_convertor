# DOC Convertor - Word文档格式化转换工具

DOC Convertor 是一个功能强大的Word文档格式化转换工具，专门用于将原始DOCX文档转换为符合特定格式要求的文档。该工具集成了多种文档处理功能，能够自动化完成文档格式化工作。

## 功能特性

### 核心转换功能
- **文档拆分**: 自动将文档拆分为封面和正文两部分
- **Pandoc转换**: 使用模板格式化正文内容
- **文档合并**: 重新合并封面和格式化后的正文
- **目录标题修改**: 自定义目录标题文本
- **图片格式化**: 图片居中显示，单倍行距
- **库号信息格式化**: 库号信息右对齐
- **分节符处理**: 在目录后插入分节符
- **页码设置**: 处理文档各节的页码设置

### 辅助功能
- **页眉页脚替换**: 替换文档的页眉页脚内容
- **表格格式化**: 表格格式标准化处理
- **文档验证**: 验证输出文档的有效性
- **中间文件保存**: 可选择保存各步骤的中间文件用于调试

## 项目结构

```
doc_convertor/
├── config/                    # 配置文件目录
│   ├── document_settings.json # 文档页面设置
│   └── picture_settings.json  # 图片格式设置
├── service/                   # 核心服务模块
│   └── converter.py           # 文档转换器主类
├── template/                  # 模板文件目录
│   ├── cover_replace_config.json # 封面替换配置
│   ├── reference_content.docx    # 正文内容模板
│   ├── reference_cover1.docx     # 封面模板1
│   └── reference_cover2.docx     # 封面模板2
├── utils/                     # 工具模块目录
│   ├── cover_replace.py          # 封面内容替换工具
│   ├── document_page_settings.py # 文档页面设置工具
│   ├── docx_header_footer_replace.py # 页眉页脚替换工具
│   ├── docx_merge.py             # 文档合并工具
│   ├── docx_picture.py           # 图片格式化工具
│   ├── docx_section_break.py     # 分节符处理工具
│   ├── docx_split.py             # 文档拆分工具
│   ├── docx_supplement.py        # 补充处理工具
│   ├── docx_table_format.py      # 表格格式化工具
│   ├── docx_update_toc_title.py  # 目录标题更新工具
│   └── pandoc_converter.py       # Pandoc转换工具
├── main.py                   # 主程序入口
└── README.md                 # 项目说明文档
```

## 安装与配置

### 环境要求
- Python 3.6+
- Windows操作系统（部分功能依赖COM组件）
- Pandoc（需要单独安装）

### 依赖库安装
```bash
pip install python-docx
pip install win32com
```

### Pandoc安装
Pandoc是一个文档转换工具，本项目依赖它来完成DOCX文档的格式化转换。

1. 访问Pandoc官方下载页面：[https://github.com/jgm/pandoc/releases](https://github.com/jgm/pandoc/releases)
2. 根据您的操作系统选择合适的安装包下载
3. Windows用户建议下载 `pandoc-x.xx-windows-x86_64.msi` 安装包
4. 运行安装程序，按照提示完成安装
5. 安装完成后，确保pandoc命令在命令行中可用：
   ```bash
   pandoc --version
   ```

或者，您也可以将pandoc.exe直接放置在系统的PATH环境变量包含的目录中，例如：
- `C:\Program Files\Pandoc\`
- `C:\Users\[用户名]\AppData\Local\Programs\Python\Python3x\Scripts\`

## 使用方法

### 基本使用
1. 修改 [main.py](file:///c%3A/Users/yanha/Desktop/%E6%95%B0%E5%AD%97%E6%80%BB%E5%B8%88/doc_convertor/main.py) 中的源文件路径
2. 确保已正确安装Pandoc
3. 运行主程序:
   ```bash
   python main.py
   ```

### 高级配置
- **模板文件**: 在 [template/](file:///c%3A/Users/yanha/Desktop/%E6%95%B0%E5%AD%97%E6%80%BB%E5%B8%88/doc_convertor/template) 目录中自定义参考文档模板
- **配置文件**: 在 [config/](file:///c%3A/Users/yanha/Desktop/%E6%95%B0%E5%AD%97%E6%80%BB%E5%B8%88/doc_convertor/config/) 目录中调整文档和图片格式设置
- **封面替换**: 在 [template/cover_replace_config.json](file:///c%3A/Users/yanha/Desktop/%E6%95%B0%E5%AD%97%E6%80%BB%E5%B8%88/doc_convertor/template/cover_replace_config.json) 中配置封面内容替换规则

### API调用
```python
from service.converter import DocumentConverter

# 创建转换器实例
with DocumentConverter() as converter:
    success = converter.convert_document(
        source_file="source.docx",
        output_file="output.docx",
        header_text="文档标题",
        toc_title="目 录",
        save_intermediate=True
    )
```

## 转换流程

1. **文档拆分**: 将源文档分离为封面和正文内容
2. **Pandoc转换**: 使用模板格式化正文内容
3. **文档合并**: 重新合并封面和格式化后的正文
4. **目录标题更新**: 修改目录标题为指定文本
5. **图片格式化**: 设置图片居中显示和行距
6. **补充处理**: 格式化库号信息等特殊内容
7. **分节符处理**: 在目录后插入分节符
8. **页码设置**: 处理各节的页码格式

## 配置说明

### document_settings.json
```json
{
  "page_settings": {
    "paper_size": {
      "width": 21.0,
      "height": 29.7,
      "unit": "cm"
    },
    "margins": {
      "top": 3.1,
      "bottom": 2.8,
      "left": 2.8,
      "right": 2.8,
      "header": 2.4,
      "footer": 2.4,
      "unit": "cm"
    }
  }
}
```

### picture_settings.json
```json
{
  "picture_format": {
    "alignment": "center",
    "line_spacing": 1.0,
    "wrap_type": "inline"
  }
}
```

## 注意事项

1. 项目主要在Windows环境下开发和测试
2. 部分功能依赖Microsoft Word COM组件
3. 建议使用绝对路径指定文件位置
4. 转换过程中会生成临时文件，完成后自动清理
5. 如需调试，可启用中间文件保存功能
6. 必须正确安装Pandoc才能使用文档格式化功能

## 贡献

欢迎提交Issue和Pull Request来改进这个工具。

## 许可证

[MIT License](LICENSE)