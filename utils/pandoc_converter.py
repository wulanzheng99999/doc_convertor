#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pandoc文档转换器
支持各种文档格式转换，特别优化了DOCX模板转换功能
"""

import os
import subprocess
import sys
from pathlib import Path
from typing import Optional, List, Dict, Any

class PandocConverter:
    """Pandoc文档转换器类"""
    
    def __init__(self, pandoc_path: Optional[str] = None):
        """
        初始化转换器
        
        Args:
            pandoc_path: pandoc.exe的路径，如果为None则使用当前目录下的pandoc.exe
        """
        if pandoc_path is None:
            # 使用当前脚本所在目录下的pandoc.exe
            current_dir = Path(__file__).parent
            self.pandoc_path = current_dir / "pandoc.exe"
        else:
            self.pandoc_path = Path(pandoc_path)
        
        # 不在初始化时检查pandoc是否存在，而是在实际使用时检查
        # 这样可以避免在DocumentConverter初始化时就失败
        print(f"Pandoc路径: {self.pandoc_path}")
        
    def _check_pandoc_available(self) -> bool:
        """
        检查Pandoc是否可用
        
        Returns:
            bool: Pandoc是否可用
        """
        try:
            if not self.pandoc_path.exists():
                return False
            
            # 尝试执行pandoc --version来验证是否可用
            import subprocess
            result = subprocess.run(
                [str(self.pandoc_path), "--version"],
                capture_output=True,
                text=True,
                timeout=10
            )
            return result.returncode == 0
        except Exception:
            return False
    
    def run_pandoc(self, args: List[str]) -> subprocess.CompletedProcess:
        """
        运行pandoc命令
        
        Args:
            args: pandoc命令参数列表
            
        Returns:
            subprocess.CompletedProcess对象
        """
        # 检查pandoc是否可用
        if not self._check_pandoc_available():
            raise FileNotFoundError(f"Pandoc不可用: {self.pandoc_path}")
        
        cmd = [str(self.pandoc_path)] + args
        print(f"执行命令: {' '.join(cmd)}")
        
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                check=True
            )
            print("转换成功!")
            return result
        except subprocess.CalledProcessError as e:
            print(f"Pandoc执行失败: {e}")
            print(f"错误输出: {e.stderr}")
            raise
        except Exception as e:
            print(f"执行过程中发生错误: {e}")
            raise
    
    def convert_with_template(
        self, 
        input_file: str, 
        output_file: str, 
        template_file: str,
        additional_args: Optional[List[str]] = None
    ) -> bool:
        """
        使用DOCX模板进行转换（保持格式、页眉页脚等）
        
        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径
            template_file: 模板DOCX文件路径
            additional_args: 额外的pandoc参数
            
        Returns:
            转换是否成功
        """
        args = [
            input_file,
            "-o", output_file,
            "--reference-doc", template_file
        ]
        
        if additional_args:
            args.extend(additional_args)
        
        try:
            self.run_pandoc(args)
            return True
        except Exception as e:
            print(f"模板转换失败: {e}")
            return False
    
    def convert_basic(
        self, 
        input_file: str, 
        output_file: str,
        from_format: Optional[str] = None,
        to_format: Optional[str] = None,
        additional_args: Optional[List[str]] = None
    ) -> bool:
        """
        基本文档转换
        
        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径
            from_format: 源格式（如果不指定，pandoc会自动检测）
            to_format: 目标格式（如果不指定，从输出文件扩展名推断）
            additional_args: 额外的pandoc参数
            
        Returns:
            转换是否成功
        """
        args = []
        
        if from_format:
            args.extend(["-f", from_format])
        
        if to_format:
            args.extend(["-t", to_format])
        
        args.extend([input_file, "-o", output_file])
        
        if additional_args:
            args.extend(additional_args)
        
        try:
            self.run_pandoc(args)
            return True
        except Exception as e:
            print(f"基本转换失败: {e}")
            return False
    
    def get_version(self) -> str:
        """获取Pandoc版本信息"""
        # 检查pandoc是否可用
        if not self._check_pandoc_available():
            return f"Pandoc不可用: {self.pandoc_path}"
        
        try:
            result = subprocess.run(
                [str(self.pandoc_path), "--version"],
                capture_output=True,
                text=True,
                encoding='utf-8'
            )
            return result.stdout.strip()
        except Exception as e:
            return f"无法获取版本信息: {e}"
    
    def list_input_formats(self) -> List[str]:
        """列出支持的输入格式"""
        # 检查pandoc是否可用
        if not self._check_pandoc_available():
            return []
        
        try:
            result = subprocess.run(
                [str(self.pandoc_path), "--list-input-formats"],
                capture_output=True,
                text=True,
                encoding='utf-8'
            )
            return result.stdout.strip().split('\n')
        except Exception:
            return []
    
    def list_output_formats(self) -> List[str]:
        """列出支持的输出格式"""
        # 检查pandoc是否可用
        if not self._check_pandoc_available():
            return []
        
        try:
            result = subprocess.run(
                [str(self.pandoc_path), "--list-output-formats"],
                capture_output=True,
                text=True,
                encoding='utf-8'
            )
            return result.stdout.strip().split('\n')
        except Exception:
            return []
    
    def create_reference_docx(self, output_path: str) -> bool:
        """
        创建默认的参考DOCX文件，可以作为模板修改
        
        Args:
            output_path: 输出的参考文档路径
            
        Returns:
            创建是否成功
        """
        # 检查pandoc是否可用
        if not self._check_pandoc_available():
            print(f"Pandoc不可用，无法创建参考文档: {self.pandoc_path}")
            return False
        
        args = [
            "-o", output_path,
            "--print-default-data-file", "reference.docx"
        ]
        
        try:
            self.run_pandoc(args)
            print(f"默认参考文档已创建: {output_path}")
            return True
        except Exception as e:
            print(f"创建参考文档失败: {e}")
            return False
    
    def convert_with_list_formatting(
        self,
        input_file: str,
        output_file: str,
        template_file: Optional[str] = None,
        list_style: str = "arabic",
        additional_args: Optional[List[str]] = None
    ) -> bool:
        """
        转换文档并设置列表编号格式
        
        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径
            template_file: 模板DOCX文件路径（可选）
            list_style: 列表样式 ('arabic', 'alpha', 'roman', 'upper-alpha', 'upper-roman')
            additional_args: 额外的pandoc参数
            
        Returns:
            转换是否成功
        """
        args = [input_file, "-o", output_file]
        
        # 如果提供了模板文件
        if template_file:
            args.extend(["--reference-doc", template_file])
        
        # 添加列表相关的参数
        list_args = []
        
        # 根据列表样式设置相应参数
        if list_style == "arabic":
            # 默认阿拉伯数字，无需额外参数
            pass
        elif list_style == "alpha":
            # 小写字母编号
            list_args.extend(["--variable", "list-style:lower-alpha"])
        elif list_style == "upper-alpha":
            # 大写字母编号
            list_args.extend(["--variable", "list-style:upper-alpha"])
        elif list_style == "roman":
            # 小写罗马数字
            list_args.extend(["--variable", "list-style:lower-roman"])
        elif list_style == "upper-roman":
            # 大写罗马数字
            list_args.extend(["--variable", "list-style:upper-roman"])
        
        args.extend(list_args)
        
        if additional_args:
            args.extend(additional_args)
        
        try:
            self.run_pandoc(args)
            print(f"列表格式转换成功，使用样式: {list_style}")
            return True
        except Exception as e:
            print(f"列表格式转换失败: {e}")
            return False
    
    def get_list_formatting_options(self) -> Dict[str, str]:
        """
        获取可用的列表编号格式选项
        
        Returns:
            字典，包含列表样式名称和描述
        """
        return {
            "arabic": "阿拉伯数字 (1, 2, 3...)",
            "alpha": "小写字母 (a, b, c...)",
            "upper-alpha": "大写字母 (A, B, C...)",
            "roman": "小写罗马数字 (i, ii, iii...)",
            "upper-roman": "大写罗马数字 (I, II, III...)"
        }

def main():
    """主函数 - 演示如何使用转换器"""
    try:
        # 创建转换器实例
        converter = PandocConverter()
        
        # 显示版本信息
        print("Pandoc版本信息:")
        print(converter.get_version())
        print("-" * 50)
        
        # 显示支持的格式
        print("支持的输入格式:")
        input_formats = converter.list_input_formats()
        print(", ".join(input_formats[:10]) + "..." if len(input_formats) > 10 else ", ".join(input_formats))
        
        print("\n支持的输出格式:")
        output_formats = converter.list_output_formats()
        print(", ".join(output_formats[:10]) + "..." if len(output_formats) > 10 else ", ".join(output_formats))
        print("-" * 50)
        
        # 演示功能
        print("\n使用示例:")
        print("1. 创建默认参考文档模板:")
        print("   converter.create_reference_docx('template.docx')")
        
        print("\n2. 使用模板转换DOCX (保持格式、页眉页脚):")
        print("   converter.convert_with_template('input.docx', 'output.docx', 'template.docx')")
        
        print("\n3. 基本格式转换:")
        print("   converter.convert_basic('input.md', 'output.docx')")
        print("   converter.convert_basic('input.docx', 'output.pdf')")
        
        print("\n4. 指定格式转换:")
        print("   converter.convert_basic('input.txt', 'output.html', from_format='markdown', to_format='html')")
        
        print("\n5. 列表编号格式转换:")
        list_options = converter.get_list_formatting_options()
        for style, desc in list_options.items():
            print(f"   converter.convert_with_list_formatting('input.md', 'output.docx', list_style='{style}')  # {desc}")
    except Exception as e:
        print(f"程序执行失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()