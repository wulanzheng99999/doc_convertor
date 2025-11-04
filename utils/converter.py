"""
完整文档格式化转换器

实现从源文档到格式化文档的完整转换流程:
1. 文档拆分 - 分离封面和正文
2. Pandoc转换 - 使用模板格式化正文(如果没有指定，默认使用template目录下的reference.docx)
3. 表格格式化 - 使用mcp服务格式化正文的表格（还没做到这一步，现在不管）
4. 文档合并 - 重新合并为完整文档
5. 标题修改 - 修改合并后的文档的目录标题
6. 图片格式化 - 图片居中，单倍行距
7. 补充处理 - 将库号右靠齐
8. 在目录之后插入分节符
9. 处理文档节的页码设置
10. 删除文件中中所有的突出显示

# 页眉替换 - 设置指定页眉内容 // 直接替换模板文件的，现在不管

"""

import os
import sys
import tempfile
import shutil
import time
import zipfile
from datetime import datetime
from typing import Optional, Tuple
from pathlib import Path

# 导入所需的工具模块
from utils.pandoc_converter import PandocConverter
from utils.docx_split import DocxSplitProcessor
from utils.docx_merge import copy_all_to_beginning
from utils.docx_update_toc_title import update_toc_title_xml

# 添加项目路径
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

# 添加lxml导入用于去除突出显示
try:
    from lxml import etree
    LXML_AVAILABLE = True
except ImportError:
    LXML_AVAILABLE = False
    etree = None


class DocumentConverter:
    """文档格式化转换器"""

    def __init__(self, document_type: int = 1):
        """初始化转换器"""
        self.temp_dir = None
        self.pandoc_converter = None
        self.intermediate_files = {}  # 保存中间文件路径
        self.debug_output_dir = os.path.join(parent_dir, 'temp')  # 中间文件保存目录
        self.save_intermediate_files = False  # 是否保存中间文件的开关
        self.document_type = document_type  # 文档类型参数

    def __enter__(self):
        """上下文管理器入口"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口，清理临时文件"""
        self.cleanup()

    def cleanup(self):
        """清理临时文件"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def validate_input_files(self, source_file: str, template_file: str) -> bool:
        """
        验证输入文件的有效性

        Args:
            source_file: 源文档路径
            template_file: 模板文档路径

        Returns:
            bool: 文件是否有效
        """
        if not os.path.exists(source_file):
            print(f"❌ 源文档不存在: {source_file}")
            return False

        if not os.path.exists(template_file):
            print(f"❌ 模板文档不存在: {template_file}")
            return False

        # 检查文件是否为DOCX格式
        if not source_file.lower().endswith('.docx'):
            print(f"❌ 源文档不是DOCX格式: {source_file}")
            return False

        if not template_file.lower().endswith('.docx'):
            print(f"❌ 模板文档不是DOCX格式: {template_file}")
            return False

        return True

    def _save_intermediate_file(self, source_path: str, step_name: str, file_description: str = "") -> None:
        """
        保存中间文件到指定目录便于查看和调试

        Args:
            source_path: 源文件路径
            step_name: 步骤名称（如 step1_split, step2_pandoc 等）
            file_description: 文件描述（如封面, 正文内容 等）
        """
        # 如果不保存中间文件，则直接返回
        if not self.save_intermediate_files:
            return
            
        try:
            # 确保输出目录存在
            if not os.path.exists(self.debug_output_dir):
                os.makedirs(self.debug_output_dir, exist_ok=True)
                print(f"📁 创建调试输出目录: {self.debug_output_dir}")

            # 生成带时间戳的文件名
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            base_name = os.path.splitext(os.path.basename(source_path))[0]
            if file_description:
                debug_filename = f"{step_name}_{file_description}_{timestamp}.docx"
            else:
                debug_filename = f"{step_name}_{base_name}_{timestamp}.docx"

            debug_path = os.path.join(self.debug_output_dir, debug_filename)

            # 复制文件
            shutil.copy2(source_path, debug_path)

            print(f"   💾 已保存调试文件: {debug_filename}")

        except Exception as e:
            print(f"   ⚠️ 保存调试文件失败: {str(e)}")

    def step0_replace_header_footer(self, source_file: str, template_file: str) -> str:
        """
        步骤0: 页眉页脚替换 - 将源文档的页眉页脚内容替换到模板文档中

        Args:
            source_file: 源文档路径（提供页眉页脚内容）
            template_file: 模板文档路径（被替换页眉页脚内容）

        Returns:
            str: 替换页眉页脚后的模板文件路径
        """
        print("-" * 50)
        print("📑 步骤0: 页眉页脚替换")

        try:
            # 确保临时目录存在
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 生成输出文件路径
            base_name = os.path.splitext(os.path.basename(template_file))[0]
            updated_template_path = os.path.join(self.temp_dir, f"{base_name}_页眉页脚替换后.docx")

            print(f"📄 源文档: {os.path.basename(source_file)}")
            print(f"📄 模板文档: {os.path.basename(template_file)}")
            print(f"📤 更新后模板: {os.path.basename(updated_template_path)}")

            # 使用docx_header_footer_replace.py中的方法进行页眉页脚替换
            try:
                from utils.docx_header_footer_replace import replace_header_footer_content
                
                # 执行页眉页脚替换
                success = replace_header_footer_content(
                    source_docx_path=source_file,
                    target_docx_path=template_file,
                    source_section_index=2,  # 从源文档第1节提取
                    target_section_index=1,  # 替换到模板文档第1节
                    save_path=updated_template_path
                )
                
                if success and os.path.exists(updated_template_path):
                    print("✅ 页眉页脚替换成功!")
                    
                    # 保存中间文件到指定目录便于查看调试
                    if self.save_intermediate_files:
                        print(f"   更新后模板: {os.path.basename(updated_template_path)}")
                        print(f"📁 正在保存step0中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(updated_template_path, "step0_header_footer", "替换后模板")
                    
                    return updated_template_path
                else:
                    print("❌ 页眉页脚替换失败，使用原始模板文件")
                    return template_file
                    
            except Exception as replace_error:
                print(f"❌ 页眉页脚替换过程中发生错误: {str(replace_error)}")
                print("   继续使用原始模板文件")
                return template_file

        except Exception as e:
            print(f"❌ 页眉页脚替换过程中发生错误: {str(e)}")
            return template_file

    def step1_split_document(self, source_file: str) -> Tuple[Optional[str], Optional[str]]:
        """
        步骤1: 文档拆分 - 将源文档分离为封面和正文

        Args:
            source_file: 源文档路径

        Returns:
            Tuple[str, str]: (封面文件路径, 不含目录的正文内容文件路径)
            
        说明:
            - 封面文档：只包含目录之前的内容（如封面、标题等），不包含目录本身
            - 正文内容文档：包含从目录之后开始的所有内容（不包括目录）
        """
        print("-" * 50)
        print("📑 步骤1: 文档拆分")

        try:
            # 确保临时目录存在
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 生成输出文件路径
            base_name = os.path.splitext(os.path.basename(source_file))[0]
            cover_toc_path = os.path.join(self.temp_dir, f"{base_name}_封面.docx")
            content_no_toc_path = os.path.join(self.temp_dir, f"{base_name}_正文内容.docx")

            print(f"📄 源文档: {os.path.basename(source_file)}")
            print(f"📤 封面输出: {os.path.basename(cover_toc_path)}")
            print(f"📤 正文内容输出: {os.path.basename(content_no_toc_path)}")

            # 使用整合的拆分方法
            processor = DocxSplitProcessor()
            success = processor.split_document_for_conversion(
                source_path=source_file,
                cover_toc_output=cover_toc_path,
                content_no_toc_output=content_no_toc_path,
                toc_keywords=['目录', '目 录','目  录','目   录','目    录','目     录','目      录','目       录','目        录','目         录', 'Contents', 'TABLE OF CONTENTS']
            )

            if success and os.path.exists(cover_toc_path) and os.path.exists(content_no_toc_path):
                print("✅ 文档拆分成功!")
                
                # 使用cover_replace.py中的便捷函数处理封面文档
                try:
                    # 导入cover_replace模块
                    from utils.cover_replace import replace_content_in_cover_auto
                    import json
                    from datetime import datetime  # 导入datetime模块
                    
                    # 配置文件路径
                    parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                    
                    # 生成处理后的封面路径
                    processed_cover_path = os.path.join(self.temp_dir, f"{base_name}_封面_处理后.docx")
                    
                    # 使用自动选择模板和配置文件的函数处理封面
                    print("🔧 使用cover_replace_auto处理封面文档...")
                    actual_path = replace_content_in_cover_auto(
                        source_docx_path=cover_toc_path,  # 使用拆分后的封面作为源
                        output_docx_path=processed_cover_path,
                        document_type=self.document_type,  # 使用文档类型参数
                        save_file=self.save_intermediate_files  # 与convert_document的save_intermediate参数关联
                    )
                    
                    # 如果处理成功，更新cover_toc_path指向处理后的文件
                    if os.path.exists(actual_path):
                        cover_toc_path = actual_path
                        print(f"✅ 封面文档处理成功，使用处理后的文件: {os.path.basename(cover_toc_path)}")
                        
                        # 如果需要保存中间文件，也将处理后的封面文件复制到调试目录
                        if self.save_intermediate_files:
                            processed_cover_filename = f"step1_split_封面处理后_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                            processed_cover_debug_path = os.path.join(self.debug_output_dir, processed_cover_filename)
                            shutil.copy2(cover_toc_path, processed_cover_debug_path)
                            print(f"   💾 已保存处理后的封面文件: {processed_cover_filename}")
                    else:
                        print("⚠️ 封面文档处理失败，使用原始拆分文件")
                except Exception as cover_error:
                    print(f"⚠️ 封面文档处理过程中发生错误: {str(cover_error)}")
                    print("   继续使用原始拆分文件")

                # 使用cover_table_replace.py中的函数替换处理后封面中的表格
                try:
                    from utils.cover_table_replace import replace_table_after_marker
                    from datetime import datetime  # 导入datetime模块
                    
                    # 生成表格替换后的封面路径
                    table_replaced_cover_path = os.path.join(self.temp_dir, f"{base_name}_封面_表格替换后.docx")
                    
                    # 使用源文档作为表格来源，处理后的封面作为目标进行表格替换
                    print("🔧 使用cover_table_replace替换处理后封面中的表格...")
                    replaced_path = replace_table_after_marker(
                        source_path=source_file,  # 使用原始源文档作为表格来源
                        target_path=cover_toc_path,  # 使用处理后的封面作为替换目标
                        marker="各专业参加设计人员名单",  # 使用默认标记
                        save_path=table_replaced_cover_path
                    )
                    
                    # 如果替换成功，更新cover_toc_path指向表格替换后的文件
                    if os.path.exists(replaced_path):
                        cover_toc_path = replaced_path
                        print(f"✅ 封面表格替换成功，使用表格替换后的文件: {os.path.basename(cover_toc_path)}")
                        
                        # 如果需要保存中间文件，也将表格替换后的封面文件复制到调试目录
                        if self.save_intermediate_files:
                            table_replaced_cover_filename = f"step1_split_封面_表格替换后_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                            table_replaced_cover_debug_path = os.path.join(self.debug_output_dir, table_replaced_cover_filename)
                            shutil.copy2(cover_toc_path, table_replaced_cover_debug_path)
                            print(f"   💾 已保存表格替换后的封面文件: {table_replaced_cover_filename}")
                    else:
                        print("⚠️ 封面表格替换失败，使用处理后的封面文件")
                except Exception as table_error:
                    print(f"⚠️ 封面表格替换过程中发生错误: {str(table_error)}")
                    print("   继续使用处理后的封面文件")

                # 对正文内容文档中的Excel表格进行转换处理
                try:
                    from utils.docx_table_excel import convert_embedded_excels_inplace
                    from datetime import datetime  # 导入datetime模块
                    
                    # 生成处理后的正文内容路径
                    processed_content_path = os.path.join(self.temp_dir, f"{base_name}_正文内容_表格处理后.docx")
                    
                    # 使用docx_table_excel处理正文中的Excel表格
                    print("🔧 使用docx_table_excel处理正文内容中的Excel表格...")
                    try:
                        converted_count = convert_embedded_excels_inplace(
                            source_docx=content_no_toc_path,
                            output_docx=processed_content_path,
                            placeholder_when_no_pandas=True
                        )
                        
                        # 如果处理成功，更新content_no_toc_path指向处理后的文件
                        if os.path.exists(processed_content_path) and converted_count >= 0:
                            content_no_toc_path = processed_content_path
                            print(f"✅ 正文内容中的Excel表格处理成功，共转换 {converted_count} 个表格")
                            
                            # 如果需要保存中间文件，也将处理后的正文文件复制到调试目录
                            if self.save_intermediate_files:
                                processed_content_filename = f"step1_split_正文内容_表格处理后_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                                processed_content_debug_path = os.path.join(self.debug_output_dir, processed_content_filename)
                                shutil.copy2(content_no_toc_path, processed_content_debug_path)
                                print(f"   💾 已保存处理后的正文内容文件: {processed_content_filename}")
                        else:
                            print("⚠️ 正文内容中的Excel表格处理失败，使用原始正文文件")
                    except Exception as convert_error:
                        print(f"⚠️ 正文内容中的Excel表格转换失败: {str(convert_error)}")
                        print("   继续使用原始正文文件")
                except Exception as content_error:
                    print(f"⚠️ 正文内容中的Excel表格处理过程中发生错误: {str(content_error)}")
                    print("   继续使用原始正文文件")

                # 保存中间文件到指定目录便于查看调试
                if self.save_intermediate_files:
                    print(f"   封面: {os.path.basename(cover_toc_path)}")
                    print(f"   正文内容: {os.path.basename(content_no_toc_path)}")
                    print(f"📁 正在保存step1中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(cover_toc_path, "step1_split", "封面")
                    self._save_intermediate_file(content_no_toc_path, "step1_split", "正文内容")

                # 保存中间文件路径
                self.intermediate_files['cover_toc'] = cover_toc_path
                self.intermediate_files['original_content'] = content_no_toc_path

                return cover_toc_path, content_no_toc_path
            else:
                print("❌ 文档拆分失败")
                return None, None

        except Exception as e:
            print(f"❌ 文档拆分过程中发生错误: {str(e)}")
            return None, None

    def step2_pandoc_convert(self, content_file: str, template_file: str) -> Optional[str]:
        """
        步骤2: Pandoc转换 - 使用模板文件格式化正文

        Args:
            content_file: 正文内容文件路径
            template_file: 模板文件路径

        Returns:
            str: Pandoc处理后的文件路径
        """
        print("-" * 50)
        print("🔄 步骤2: Pandoc转换")

        try:
            # 初始化Pandoc转换器
            if self.pandoc_converter is None:
                # 查找pandoc可执行文件
                pandoc_path = self._find_pandoc_executable()
                if not pandoc_path:
                    print("⚠️ Pandoc不可用，跳过Pandoc转换步骤")

                    # 生成一个标记后的文件，表示跳过了Pandoc转换
                    base_name = os.path.splitext(os.path.basename(content_file))[0]
                    if not self.temp_dir:
                        raise ValueError("临时目录未初始化")
                    skipped_output = os.path.join(self.temp_dir, f"{base_name}_跳过Pandoc转换.docx")
                    shutil.copy2(content_file, skipped_output)

                    # 保存中间文件
                    self.intermediate_files['pandoc_converted'] = skipped_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(skipped_output, "step2_pandoc", "跳过转换")
                    
                    return skipped_output

                try:
                    self.pandoc_converter = PandocConverter(pandoc_path)
                except Exception as init_error:
                    print(f"⚠️ Pandoc初始化失败: {init_error}")
                    print("跳过Pandoc转换步骤")

                    # 生成一个标记后的文件，表示跳过了Pandoc转换
                    base_name = os.path.splitext(os.path.basename(content_file))[0]
                    if not self.temp_dir:
                        raise ValueError("临时目录未初始化")
                    init_failed_output = os.path.join(self.temp_dir, f"{base_name}_Pandoc初始化失败.docx")
                    shutil.copy2(content_file, init_failed_output)

                    # 保存中间文件
                    self.intermediate_files['pandoc_converted'] = init_failed_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(init_failed_output, "step2_pandoc", "初始化失败")
                    
                    return init_failed_output

            # 生成Pandoc输出文件路径
            base_name = os.path.splitext(os.path.basename(content_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            pandoc_output = os.path.join(self.temp_dir, f"{base_name}_pandoc转换.docx")

            if self.save_intermediate_files:
                print(f"📄 正在使用模板转换: {os.path.basename(template_file)}")
                print(f"📤 输出文件: {os.path.basename(pandoc_output)}")

            # 使用模板进行转换，保持表格结构
            success = self.pandoc_converter.convert_with_template(
                input_file=content_file,
                output_file=pandoc_output,
                template_file=template_file,
                additional_args=[
                    "--preserve-tabs",           # 保持制表符
                    "--wrap=none",              # 不自动换行
                    "--reference-links",        # 使用引用链接
                    "--columns=80",     # 设置合适的列宽
                    "--table-of-contents",  # 保持目录结构
                    "--standalone",      # 独立文档模式
                ]
            )

            if success and os.path.exists(pandoc_output):
                print("✅ Pandoc转换成功!")
                # 保存中间文件
                self.intermediate_files['pandoc_converted'] = pandoc_output

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"   转换后文件: {os.path.basename(pandoc_output)}")
                    print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(pandoc_output, "step2_pandoc", "转换成功")

                return pandoc_output
            else:
                print("❌ Pandoc转换失败，使用原始文件继续")

                # 在转换失败时，复制原文件作为备用
                base_name = os.path.splitext(os.path.basename(content_file))[0]
                fallback_output = os.path.join(self.temp_dir, f"{base_name}_Pandoc失败备用.docx")
                shutil.copy2(content_file, fallback_output)

                # 保存中间文件
                self.intermediate_files['pandoc_converted'] = fallback_output

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(fallback_output, "step2_pandoc", "失败备用")
                
                return fallback_output

        except Exception as e:
            print(f"❌ Pandoc转换过程中发生错误: {str(e)}")
            print("使用原始文件继续后续处理")

            # 在发生异常时，复制原文件作为备用
            base_name = os.path.splitext(os.path.basename(content_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            error_output = os.path.join(self.temp_dir, f"{base_name}_Pandoc异常备用.docx")
            shutil.copy2(content_file, error_output)

            # 保存中间文件
            self.intermediate_files['pandoc_converted'] = error_output

            # 保存调试文件
            if self.save_intermediate_files:
                print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                self._save_intermediate_file(error_output, "step2_pandoc", "异常备用")
            
            return error_output

    def step3_format_tables(self, content_file: str, template_file: str, original_content_file: Optional[str] = None) -> Optional[str]:
        """
        步骤3: 表格格式化 - 使用原始内容文件中的表格替换处理后的文件中的表格

        Args:
            content_file: 正文内容文件路径（被替换表格内容）
            template_file: 模板文件路径
            original_content_file: 原始正文内容文件路径（提供表格内容），如果提供则使用表格替换功能

        Returns:
            str: 表格格式化后的文件路径
        """
        print("-" * 50)
        print("📊 步骤3: 表格格式化")

        try:
            # 确保临时目录存在
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 生成表格格式化输出文件路径
            base_name = os.path.splitext(os.path.basename(content_file))[0]
            
            # 如果提供了原始内容文件，则使用表格替换功能
            if original_content_file and os.path.exists(original_content_file):
                formatted_output = os.path.join(self.temp_dir, f"{base_name}_表格替换.docx")
                print("🔧 使用表格替换功能处理表格...")
                
                # 使用docx_table_replace.py中的方法进行表格替换
                try:
                    from utils.docx_table_replace import replace_tables_by_index
                    
                    # 使用原始正文内容文件中的表格替换Pandoc转换后的文件中的表格
                    success = replace_tables_by_index(
                        original_path=original_content_file,  # 原始正文内容文件（提供表格）
                        edited_path=content_file,             # Pandoc转换后的文件（被替换表格）
                        output_path=formatted_output          # 输出文件
                    )
                    
                    if success and os.path.exists(formatted_output):
                        print("✅ 表格替换成功!")
                        
                        # 保存中间文件
                        self.intermediate_files['table_replaced'] = formatted_output

                        # 保存调试文件
                        if self.save_intermediate_files:
                            print(f"   表格替换后文件: {os.path.basename(formatted_output)}")
                            print(f"📁 正在保存step3中间文件到: {self.debug_output_dir}")
                            self._save_intermediate_file(formatted_output, "step3_table", "替换完成")
                        
                        return formatted_output
                    else:
                        print("❌ 表格替换失败，使用原始文件继续")
                        return content_file
                        
                except Exception as replace_error:
                    print(f"❌ 表格替换过程中发生错误: {str(replace_error)}")
                    print("使用原始文件继续后续处理")
                    return content_file
            else:
                print("⚠️ 未提供原始内容文件，跳过表格替换步骤")
                return content_file

        except Exception as e:
            print(f"❌ 表格格式化过程中发生错误: {str(e)}")
            return content_file

    def step4_merge_documents(self, cover_toc_file: str, processed_content_file: str, output_file: str) -> bool:
        """
        步骤4: 文档合并 - 将封面添加到正文开始

        Args:
            cover_toc_file: 封面文件路径
            processed_content_file: 处理后的正文文件路径
            output_file: 最终输出文件路径

        Returns:
            bool: 合并是否成功
        """
        print("-" * 50)
        print("📚 步骤4: 文档合并")

        try:
            if self.save_intermediate_files:
                print(f"📄 封面: {os.path.basename(cover_toc_file)}")
                print(f"📄 正文内容: {os.path.basename(processed_content_file)}")
                print(f"📤 最终输出: {os.path.basename(output_file)}")

            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                print(f"📁 创建输出目录: {output_dir}")

            # 使用docx_merge.py中的方法进行文档合并
            try:
                copy_all_to_beginning(
                    file_a=cover_toc_file,
                    file_b=processed_content_file,
                    output_file=output_file
                )
                
                if os.path.exists(output_file):
                    print("✅ 文档合并成功!")
                    
                    # 验证输出文件的有效性
                    if self._validate_output_file(output_file):
                        print("✅ 输出文件验证通过")
                        # 保存最终文件
                        self.intermediate_files['final_document'] = output_file

                        # 保存调试文件
                        if self.save_intermediate_files:
                            print(f"   最终文档: {os.path.basename(output_file)}")
                            print(f"📁 正在保存step4最终文件到: {self.debug_output_dir}")
                            self._save_intermediate_file(output_file, "step4_final", "最终文档")
                        return True
                    else:
                        print("❌ 输出文件验证失败")
                        return False
                    
                else:
                    print("❌ 文档合并失败，输出文件不存在")
                    return False
                    
            except Exception as merge_error:
                print(f"❌ 使用docx_merge方法合并文档时发生错误: {str(merge_error)}")
                return False

        except Exception as e:
            print(f"❌ 文档合并过程中发生错误: {str(e)}")
            return False

    def step5_update_toc_title(self, docx_file: str, new_title: str = "目录") -> bool:
        """
        步骤5: 更新目录标题 - 修改文档中的目录标题

        Args:
            docx_file: 需要修改的文档路径
            new_title: 新的目录标题

        Returns:
            bool: 更新是否成功
        """
        print("-" * 50)
        print("🏷️ 步骤5: 更新目录标题")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")
                print(f"🔤 新标题: '{new_title}'")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 使用docx_update_toc_title中的方法更新目录标题
            try:
                # 使用XML方式更新目录标题，保留原有格式
                update_toc_title_xml(docx_file, new_title)
                print("✅ 目录标题更新成功!")
                return True
            except Exception as xml_error:
                print(f"⚠️ XML方式更新目录标题失败: {str(xml_error)}")
                print("尝试使用COM方式更新目录标题...")
                
                try:
                    # 使用COM方式更新目录标题
                    from utils.docx_update_toc_title import update_toc_title
                    update_toc_title(docx_file, new_title)
                    print("✅ 目录标题更新成功!")
                    return True
                except Exception as com_error:
                    print(f"❌ COM方式更新目录标题也失败: {str(com_error)}")
                    return False

        except Exception as e:
            print(f"❌ 更新目录标题过程中发生错误: {str(e)}")
            return False

    def step6_format_pictures(self, docx_file: str) -> bool:
        """
        步骤6: 图片格式化 - 图片居中，单倍行距
        
        Args:
            docx_file: 需要处理的文档路径
            
        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("🖼️ 步骤6: 图片格式化")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 生成处理后的文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_图片格式化.docx")

            # 使用docx_picture.py中的高级处理方式
            try:
                from utils.docx_picture import format_pictures_with_advanced_settings
                
                # 调用高级图片格式化函数
                success = format_pictures_with_advanced_settings(
                    doc_path=docx_file,
                    save_path=formatted_output
                )
                
                if success and os.path.exists(formatted_output):
                    print("✅ 图片格式化成功!")
                    
                    # 保存中间文件
                    self.intermediate_files['picture_formatted'] = formatted_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   图片格式化后文件: {os.path.basename(formatted_output)}")
                        print(f"📁 正在保存step6中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(formatted_output, "step6_picture", "格式化完成")
                    
                    # 将处理后的文件复制回原文件路径，以便后续步骤使用
                    shutil.copy2(formatted_output, docx_file)
                    return True
                else:
                    print("❌ 图片格式化失败")
                    return False
                    
            except Exception as format_error:
                print(f"❌ 图片格式化过程中发生错误: {str(format_error)}")
                import traceback
                traceback.print_exc()
                return False

        except Exception as e:
            print(f"❌ 图片格式化过程中发生错误: {str(e)}")
            return False

    def step7_format_library_number(self, docx_file: str) -> bool:
        """
        步骤7: 库号信息格式化 - 将库号信息靠右对齐
        
        Args:
            docx_file: 需要处理的文档路径
            
        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("🔢 步骤7: 库号信息格式化")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 生成处理后的文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_库号格式化.docx")

            # 使用docx_supplement.py中的高级处理方式
            try:
                from utils.docx_supplement import format_library_number_advanced
                
                # 调用高级库号格式化函数
                success = format_library_number_advanced(
                    doc_path=docx_file,
                    save_path=formatted_output
                )
                
                if success and os.path.exists(formatted_output):
                    print("✅ 库号信息格式化成功!")
                    
                    # 保存中间文件
                    self.intermediate_files['library_number_formatted'] = formatted_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   库号格式化后文件: {os.path.basename(formatted_output)}")
                        print(f"📁 正在保存step7中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(formatted_output, "step7_library_number", "格式化完成")
                    
                    # 将处理后的文件复制回原文件路径，以便后续步骤使用
                    shutil.copy2(formatted_output, docx_file)
                    return True
                else:
                    print("❌ 库号信息格式化失败")
                    return False
                    
            except Exception as format_error:
                print(f"❌ 库号信息格式化过程中发生错误: {str(format_error)}")
                import traceback
                traceback.print_exc()
                return False

        except Exception as e:
            print(f"❌ 库号信息格式化过程中发生错误: {str(e)}")
            return False

    def step8_insert_section_break(self, docx_file: str) -> bool:
        """
        步骤8: 在目录后插入分节符
        
        Args:
            docx_file: 需要处理的文档路径
            
        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("📑 步骤8: 在目录后插入分节符")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 添加延迟，确保之前的COM操作完全释放资源
            print("⏳ 等待COM资源释放...")
            import time
            time.sleep(3)

            # 生成处理后的文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_插入分节符.docx")

            # 使用docx_supplement.py中的方法
            try:
                from utils.docx_supplement import insert_section_break_after_toc
                
                # 调用插入分节符函数
                success = insert_section_break_after_toc(
                    doc_path=docx_file,
                    save_path=formatted_output
                )
                
                if success and os.path.exists(formatted_output):
                    print("✅ 在目录后插入分节符成功!")
                    
                    # 保存中间文件
                    self.intermediate_files['section_break_inserted'] = formatted_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   插入分节符后文件: {os.path.basename(formatted_output)}")
                        print(f"📁 正在保存step8中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(formatted_output, "step8_section_break", "插入完成")
                    
                    # 将处理后的文件复制回原文件路径，以便后续步骤使用
                    shutil.copy2(formatted_output, docx_file)
                    return True
                else:
                    print("❌ 在目录后插入分节符失败")
                    # 尝试使用备选方法
                    print("🔄 尝试使用备选方法...")
                    return self._insert_section_break_fallback(docx_file, formatted_output)
                    
            except Exception as format_error:
                print(f"❌ 在目录后插入分节符过程中发生错误: {str(format_error)}")
                # 尝试使用备选方法
                print("🔄 尝试使用备选方法...")
                return self._insert_section_break_fallback(docx_file, formatted_output)

        except Exception as e:
            print(f"❌ 在目录后插入分节符过程中发生错误: {str(e)}")
            return False

    def step9_process_sections(self, docx_file: str) -> bool:
        """
        步骤9: 处理文档节的页码设置
        - 取消第三节与第二节的链接
        - 处理第二节的页码（删除PAGE域）
        - 处理第三节的页码（重置为1）
        
        Args:
            docx_file: 需要处理的文档路径
            
        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("📑 步骤9: 处理文档节的页码设置")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 添加延迟，确保之前的COM操作完全释放资源
            print("⏳ 等待COM资源释放...")
            import time
            time.sleep(3)

            # 生成临时文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            
            temp_file1 = os.path.join(self.temp_dir, f"{base_name}_取消节链接.docx")
            temp_file2 = os.path.join(self.temp_dir, f"{base_name}_处理第二节页码.docx")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_处理第三节页码.docx")

            # 步骤9.1: 使用docx_supplement.py中的方法取消第三节与第二节的链接
            print("\n步骤9.1: 取消第三节与第二节的链接...")
            try:
                from utils.docx_supplement import cancel_section_link_com
                
                success = cancel_section_link_com(
                    doc_path=docx_file,
                    save_path=temp_file1,
                    section_number=3  # 第三节
                )
                
                if not success or not os.path.exists(temp_file1):
                    print("❌ 步骤9.1失败，无法继续执行后续步骤")
                    return False
                else:
                    print("✅ 步骤9.1完成")
                    
                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   取消节链接后文件: {os.path.basename(temp_file1)}")
                        print(f"📁 正在保存step9.1中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(temp_file1, "step9_section_link", "取消链接完成")
            except Exception as e:
                print(f"❌ 步骤9.1失败: {e}")
                return False

            # 步骤9.2: 使用docx_supplement.py中的方法处理第二节的页码
            print("\n步骤9.2: 处理第二节的页码...")
            try:
                from utils.docx_supplement import process_section2_docx
                
                process_section2_docx(temp_file1, temp_file2, section_index=2)
                print("✅ 步骤9.2完成")
                
                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"   处理第二节页码后文件: {os.path.basename(temp_file2)}")
                    print(f"📁 正在保存step9.2中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(temp_file2, "step9_section2_page", "处理完成")
            except Exception as e:
                print(f"❌ 步骤9.2失败: {e}")
                return False

            # 步骤9.3: 使用docx_supplement.py中的方法处理第三节的页码
            print("\n步骤9.3: 处理第三节的页码...")
            try:
                from utils.docx_supplement import process_section3_docx
                
                process_section3_docx(temp_file2, formatted_output)
                print("✅ 步骤9.3完成")
                
                # 保存中间文件
                self.intermediate_files['section_page_processed'] = formatted_output

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"   处理第三节页码后文件: {os.path.basename(formatted_output)}")
                    print(f"📁 正在保存step9.3中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(formatted_output, "step9_section3_page", "处理完成")
                
                # 将处理后的文件复制回原文件路径，以便后续步骤使用
                shutil.copy2(formatted_output, docx_file)
                return True
            except Exception as e:
                print(f"❌ 步骤9.3失败: {e}")
                return False

        except Exception as e:
            print(f"❌ 处理文档节的页码设置过程中发生错误: {str(e)}")
            return False

    def _insert_section_break_fallback(self, docx_file: str, formatted_output: str) -> bool:
        """
        备选方法：在目录后插入分节符的降级处理
        
        Args:
            docx_file: 需要处理的文档路径
            formatted_output: 输出文件路径
            
        Returns:
            bool: 处理是否成功
        """
        try:
            print("🔧 尝试使用XML方法插入分节符...")
            
            # 使用XML方法作为备选
            from utils.docx_section_break import insert_section_break_after_toc_xml
            
            success = insert_section_break_after_toc_xml(
                doc_path=docx_file,
                save_path=formatted_output
            )
            
            if success and os.path.exists(formatted_output):
                print("✅ 使用XML方法插入分节符成功!")
                return True
            else:
                print("❌ 使用XML方法插入分节符失败")
                return False
                
        except Exception as xml_error:
            print(f"❌ 使用XML方法插入分节符时发生错误: {str(xml_error)}")
            print("⚠️ 无法在目录后插入分节符，继续使用原有格式")
            return False

    def _validate_output_file(self, file_path: str) -> bool:
        """
        验证输出文件的有效性

        Args:
            file_path: 文件路径

        Returns:
            bool: 文件是否有效
        """
        try:
            if not os.path.exists(file_path):
                return False

            # 检查文件大小
            file_size = os.path.getsize(file_path)
            if file_size < 1000:  # 小于1KB可能有问题
                print(f"⚠️ 文件大小异常: {file_size} bytes")
                return False

            # 检查是否为有效的ZIP文件（DOCX本质上ZIP文件）
            import zipfile
            try:
                with zipfile.ZipFile(file_path, 'r') as zip_file:
                    # 检查必要的文件
                    required_files = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml']
                    file_list = zip_file.namelist()

                    for req_file in required_files:
                        if req_file not in file_list:
                            print(f"⚠️ 缺少必要文件: {req_file}")
                            return False

                    return True
            except zipfile.BadZipFile:
                print("⚠️ 文件不是有效的ZIP格式")
                return False

        except Exception as e:
            print(f"⚠️ 文件验证过程中发生错误: {str(e)}")
            return False

    def step10_remove_highlights(self, source_file: str) -> str:
        """
        步骤10: 删除文档中所有的突出显示（高亮、底纹、颜色）
        
        Args:
            source_file: 源文档路径
            
        Returns:
            str: 去除突出显示后的文档路径
        """
        print("-" * 50)
        print("📑 步骤10: 删除文档中所有的突出显示")

        try:
            # 确保临时目录存在
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 检查lxml是否可用
            if not LXML_AVAILABLE:
                print("⚠️ lxml库不可用，跳过突出显示删除步骤")
                return source_file

            # 生成输出文件路径
            base_name = os.path.splitext(os.path.basename(source_file))[0]
            no_highlight_file = os.path.join(self.temp_dir, f"{base_name}_无突出显示.docx")

            print(f"📄 输入文档: {os.path.basename(source_file)}")
            print(f"📤 输出文档: {os.path.basename(no_highlight_file)}")

            # 使用lxml删除Word文件中的所有高亮
            try:
                # 定义命名空间和常量
                W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                NSMAP = {"w": W_NS}
                XML_PARSER = etree.XMLParser(ns_clean=True, recover=True, remove_blank_text=False) if (LXML_AVAILABLE and etree is not None) else None
                REMOVE_COLOR_NODE = True  # 彻底移除颜色节点

                def process_xml_bytes(data: bytes, remove_color_node=False) -> bytes:
                    """删除 w:highlight、w:shd，并处理 w:color"""
                    if not LXML_AVAILABLE or etree is None:
                        return data
                    
                    try:
                        root = etree.fromstring(data, parser=XML_PARSER) if etree is not None else None
                        if root is None:
                            return data
                    except Exception:
                        return data

                    changed = False
                    # 删除 highlight
                    for node in root.xpath('.//w:highlight', namespaces=NSMAP):
                        parent = node.getparent()
                        if parent is not None:
                            parent.remove(node)
                            changed = True

                    # 删除底纹
                    for node in root.xpath('.//w:shd', namespaces=NSMAP):
                        parent = node.getparent()
                        if parent is not None:
                            parent.remove(node)
                            changed = True

                    # 处理颜色
                    for color in root.xpath('.//w:color', namespaces=NSMAP):
                        val = color.get("val")
                        if val is not None and val.lower() != "auto":
                            if remove_color_node:
                                parent = color.getparent()
                                if parent is not None:
                                    parent.remove(color)
                                    changed = True
                            else:
                                color.set("val", "auto")
                                changed = True

                    if not changed:
                        return data
                    return etree.tostring(root, encoding="utf-8", xml_declaration=True) if (LXML_AVAILABLE and etree is not None) else data

                # 处理DOCX文件
                src = Path(source_file)
                if not src.exists():
                    raise FileNotFoundError(f"文件不存在: {src}")

                dest = Path(no_highlight_file)

                with zipfile.ZipFile(src, 'r') as zin:
                    with zipfile.ZipFile(dest, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                        for name in zin.namelist():
                            data = zin.read(name)
                            if name.startswith("word/") and name.endswith(".xml"):
                                try:
                                    new_data = process_xml_bytes(data, remove_color_node=REMOVE_COLOR_NODE)
                                    zout.writestr(name, new_data)
                                except Exception as e:
                                    print(f"⚠ 处理 {name} 出错，保留原文件。错误：{e}")
                                    zout.writestr(name, data)
                            else:
                                zout.writestr(name, data)

                print("✅ 突出显示删除成功!")

                # 保存中间文件到指定目录便于查看调试
                if self.save_intermediate_files:
                    print(f"   无突出显示文档: {os.path.basename(no_highlight_file)}")
                    print(f"📁 正在保存step10中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(no_highlight_file, "step10_highlights", "无突出显示")

                # 保存中间文件路径
                self.intermediate_files['no_highlights'] = no_highlight_file

                return no_highlight_file

            except Exception as remove_error:
                print(f"❌ 突出显示删除过程中发生错误: {str(remove_error)}")
                print("   继续使用原始文档")
                return source_file

        except Exception as e:
            print(f"❌ 突出显示删除过程中发生错误: {str(e)}")
            return source_file

    def _find_pandoc_executable(self) -> Optional[str]:
        """
        查找Pandoc可执行文件

        Returns:
            str: Pandoc可执行文件路径
        """
        import subprocess

        # 可能的Pandoc位置 - 优先utils目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(current_dir)
        utils_pandoc = os.path.join(parent_dir, 'utils', 'pandoc.exe')

        possible_paths = [
            # 优先使用utils目录中的pandoc.exe
            utils_pandoc,
            # 系统PATH中的pandoc
            "pandoc",
            "pandoc.exe",
            # 当前目录下的pandoc.exe
            os.path.join(os.path.dirname(__file__), "pandoc.exe"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "pandoc.exe"),
            # 常见安装位置
            r"C:\Program Files\Pandoc\pandoc.exe",
            r"C:\Program Files (x86)\Pandoc\pandoc.exe",
            # Conda环境
            os.path.join(os.environ.get('CONDA_PREFIX', ''), 'Scripts', 'pandoc.exe'),
            os.path.join(os.environ.get('CONDA_PREFIX', ''), 'bin', 'pandoc')
        ]

        for path in possible_paths:
            if not path:  # 跳过空路径
                continue

            try:
                # 尝试执行pandoc --version
                result = subprocess.run(
                    [path, "--version"],
                    capture_output=True,
                    text=True,
                    timeout=10,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )
                if result.returncode == 0:
                    return path
            except (subprocess.TimeoutExpired, FileNotFoundError, PermissionError, Exception):
                continue

        print("⚠️ 未找到Pandoc可执行文件")
        print("请通过以下方式安装Pandoc:")
        print("1. 下载安装: https://pandoc.org/installing.html")
        print("2. 使用conda: conda install pandoc")
        print("3. 使用choco: choco install pandoc")
        print(f"4. utils目录路径: {utils_pandoc}")
        return None

    def convert_document(
            self,
            source_file: str,
            output_file: str,
            template_file: Optional[str] = None,
            header_text: str = "格式化文档",
            toc_title: str = "目 录",
            save_intermediate: bool = False,
            intermediate_dir: Optional[str] = None,
            document_type: int = 1
    ) -> bool:
        """
        完整的文档格式化转换流程

        Args:
            source_file: 源文档路径
            output_file: 输出文档路径
            template_file: 模板文档路径（可选，默认使用template/reference.docx）
            header_text: 页眉文本
            toc_title: 目录标题（可选，默认为"目 录"）
            save_intermediate: 是否保存中间文件（默认为False）
            intermediate_dir: 中间文件保存目录（仅在save_intermediate为True时有效）
            document_type: 文档类型 (1, 2, 3, 4)

        Returns:
            bool: 转换是否成功
        """

        # 设置是否保存中间文件
        self.save_intermediate_files = save_intermediate
        
        # 设置文档类型
        self.document_type = document_type
        print(f"   文档类型: {document_type}")
        # 如果指定了中间文件目录，则使用该目录
        if save_intermediate and intermediate_dir:
            self.debug_output_dir = intermediate_dir
        
        # 如果未指定模板文件，则使用默认模板
        if template_file is None:
            # 获取项目根目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            parent_dir = os.path.dirname(current_dir)
            template_file = os.path.join(parent_dir, 'template', 'reference_content.docx')
            if self.save_intermediate_files:
                print(f"信息: 未指定模板文件，使用默认模板: {template_file}")

        start_time = time.time()

        print("🚀 开始文档格式化转换")
        print("=" * 80)
        print(f"📁 源文档: {source_file}")
        print(f"📄 模板文档: {template_file}")
        print(f"📤 输出文档: {output_file}")
        print(f"📋 页眉文本: '{header_text}'")
        print(f"📋 目录标题: '{toc_title}'")
        if save_intermediate:
            print(f"💾 保存中间文件: 是")
            print(f"📂 中间文件目录: {self.debug_output_dir}")
        else:
            print(f"💾 保存中间文件: 否")
        print(f"⏰ 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 80)
        print()

        # 验证输入文件
        if not self.validate_input_files(source_file, template_file):
            return False

        # 创建临时目录
        self.temp_dir = tempfile.mkdtemp(prefix="doc_converter_")
        if self.save_intermediate_files:
            print(f"📁 临时目录: {self.temp_dir}")

        try:
            # 步骤0: 页眉页脚替换
            updated_template_file = self.step0_replace_header_footer(source_file, template_file)
            # 使用更新后的模板文件进行后续处理
            template_file = updated_template_file

            # 步骤1: 文档拆分
            cover_toc_file, content_file = self.step1_split_document(source_file)
            if not cover_toc_file or not content_file:
                print("❌ 转换失败: 文档拆分失败")
                return False

            # 步骤2: Pandoc转换
            pandoc_file = self.step2_pandoc_convert(content_file, template_file)
            if not pandoc_file:
                print("⚠️ Pandoc转换失败，使用原始正文文件继续")
                pandoc_file = content_file

            # 步骤3: 表格格式化
            # 使用步骤1拆分后的正文文件中的表格替换步骤2 Pandoc转换后的文件中的表格
            table_formatted_file = self.step3_format_tables(
                content_file=pandoc_file, 
                template_file=template_file,
                original_content_file=content_file  # 传入原始正文内容文件
            )
            if not table_formatted_file:
                print("⚠️ 表格格式化失败，使用Pandoc转换后的文件继续")
                table_formatted_file = pandoc_file

            # 步骤4: 文档合并
            success = self.step4_merge_documents(cover_toc_file, table_formatted_file, output_file)

            if success:
                # 步骤5: 更新目录标题
                toc_update_success = self.step5_update_toc_title(output_file, toc_title)
                # if toc_update_success:
                #     print("✅ 目录标题更新完成!")
                # else:
                #     print("⚠️ 目录标题更新失败，继续使用原有标题")
                if not toc_update_success:
                    print("⚠️ 目录标题更新失败，继续使用原有标题")

                # 步骤6: 图片格式化
                picture_format_success = self.step6_format_pictures(output_file)
                if not picture_format_success:
                    print("⚠️ 图片格式化失败，继续使用原有格式")

                # 步骤7: 库号信息格式化
                library_number_format_success = self.step7_format_library_number(output_file)
                if not library_number_format_success:
                    print("⚠️ 库号信息格式化失败，继续使用原有格式")

                # 步骤8: 在目录后插入分节符
                section_break_insert_success = self.step8_insert_section_break(output_file)
                if not section_break_insert_success:
                    print("⚠️ 在目录后插入分节符失败，继续使用原有格式")

                # 步骤9: 处理文档节的页码设置
                section_page_process_success = self.step9_process_sections(output_file)
                if not section_page_process_success:
                    print("⚠️ 处理文档节的页码设置失败，继续使用原有格式")

                # 步骤10: 删除文档中所有的突出显示
                no_highlights_file = self.step10_remove_highlights(output_file)
                # 如果成功去除突出显示，将结果复制回输出文件
                if no_highlights_file != output_file and os.path.exists(no_highlights_file):
                    shutil.copy2(no_highlights_file, output_file)
                    print("✅ 突出显示删除成功!")
                else:
                    print("⚠️ 删除突出显示失败，继续使用原有格式")

                end_time = time.time()
                duration = end_time - start_time

                print("\n" + "=" * 80)
                print("✅ 文档转换成功!")
                print(f"⏱️ 总耗时: {duration:.2f} 秒")
                print(f"📤 最终文档: {output_file}")
                
                # 显示中间文件信息
                if save_intermediate:
                    print("\n📋 中间文件保存在临时目录:")
                    for key, path in self.intermediate_files.items():
                        if os.path.exists(path):
                            print(f"   {key}: {os.path.basename(path)}")

                    print(f"\n📁 所有中间文件已同步保存到: {self.debug_output_dir}")
                    print("🔍 您可以在该目录中查看每个步骤的处理结果，便于调试和优化")
                else:
                    print("\n📋 中间文件未保存（根据设置）")

                print("\n💡 提示:")
                print("   - 在Word中打开文档，右键目录选择'更新域'来刷新页码")
                print("   - 检查文档格式是否符合要求")
                print("=" * 80)

                return True
            else:
                print("❌ 转换失败: 文档合并失败")
                return False

        except Exception as e:
            print(f"\n❌ 转换过程中发生错误: {str(e)}")
            return False

        finally:
            # 不立即清理临时文件，保留中间结果供调试
            if self.save_intermediate_files:
                print(f"\n📁 临时文件保留在: {self.temp_dir}")
                print("   您可以手动删除该目录，或重启程序时自动清理")


def quick_convert_document(
        source_file: str,
        output_file: str,
        template_file: Optional[str] = None,
        header_text: str = "格式化文档",
        toc_title: str = "目 录",
        save_intermediate: bool = False,
        intermediate_dir: Optional[str] = None,
        document_type: int = 1
) -> bool:
    """
    便捷函数: 快速进行文档格式化转换

    Args:
        source_file: 源文档路径
        output_file: 输出文档路径
        template_file: 模板文档路径（可选，默认使用template/reference.docx）
        header_text: 页眉文本
        toc_title: 目录标题（可选，默认为"目 录"）
        save_intermediate: 是否保存中间文件（默认为False）
        intermediate_dir: 中间文件保存目录（仅在save_intermediate为True时有效）
        document_type: 文档类型 (1, 2, 3, 4)

    Returns:
        bool: 转换是否成功
    """
    with DocumentConverter(document_type=document_type) as converter:
        return converter.convert_document(
            source_file=source_file,
            output_file=output_file,
            template_file=template_file,
            header_text=header_text,
            toc_title=toc_title,
            save_intermediate=save_intermediate,
            intermediate_dir=intermediate_dir,
            document_type=document_type
        )


if __name__ == "__main__":
    source_file = r"C:\Users\yanha\Desktop\数字总师\文档\可行性报告（test）.docx"
    
    # 创建result目录
    result_dir = os.path.join(current_dir, "result")
    os.makedirs(result_dir, exist_ok=True)
    
    # 输出文件路径
    output_file = os.path.join(result_dir, "formatted_document.docx")
    
    # 使用默认模板（template/reference.docx）
    
    # 执行转换
    with DocumentConverter(document_type=1) as converter:  # 添加文档类型参数
        success = converter.convert_document(
            source_file=source_file,
            output_file=output_file,
            header_text="数字总师可行性报告",
            toc_title="目      录",
            save_intermediate=False,
            document_type=1  # 添加文档类型参数
        )
