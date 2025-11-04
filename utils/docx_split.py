"""
DOCX 文档拆分工具
重新实现封面目录文档和不含目录的正文内容文档的生成逻辑
"""

import os
import tempfile
import shutil
import zipfile
import re
from typing import Optional, Tuple, List


class DocxSplitProcessor:
    """
    DOCX文档拆分处理工具类
    重新实现文档拆分逻辑，生成封面目录文档和不含目录的正文内容文档
    """
    
    def __init__(self):
        self.temp_dir = None
    
    def split_document_for_conversion(self, 
                                    source_path: str, 
                                    cover_toc_output: str, 
                                    content_no_toc_output: str,
                                    toc_keywords: Optional[List[str]] = None) -> bool:
        """
        将DOCX文档拆分为封面目录文档和不含目录的正文内容文档
        
        Args:
            source_path: 源文档路径
            cover_toc_output: 封面目录输出路径（不包含目录本身）
            content_no_toc_output: 不含目录的正文内容输出路径
            toc_keywords: 目录识别关键词列表
            
        Returns:
            bool: 操作是否成功
            
        说明:
            - 封面目录文档：只包含目录之前的内容（如封面、标题等），不包含目录本身
            - 正文内容文档：包含从目录之后开始的所有内容（不包括目录）
        """
        if toc_keywords is None:
            toc_keywords = ['目录', '目 录', 'Contents', 'TABLE OF CONTENTS', 'CONTENTS']
        
        try:
            # 对原始文档进行两次独立处理
            # 1. 生成封面目录文档（不包含目录本身）
            cover_success = self._process_cover_document(source_path, cover_toc_output, toc_keywords)
            
            # 2. 生成不含目录的正文内容文档
            content_success = self._process_content_no_toc_document(source_path, content_no_toc_output, toc_keywords)
            
            if cover_success and content_success:
                return True
            else:
                print(f"❌ 文档拆分失败: 封面文档处理={cover_success}, 正文文档处理={content_success}")
                return False
            
        except Exception as e:
            print(f"❌ 文档拆分失败: {str(e)}")
            return False
    
    def _process_cover_document(self, source_path: str, cover_toc_output: str, toc_keywords: List[str]) -> bool:
        """
        处理生成封面目录文档（不包含目录本身）
        使用与docx_split_merge.py中相同的逻辑
        """
        temp_dir = None
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()
            source_temp = os.path.join(temp_dir, 'source')
            cover_temp = os.path.join(temp_dir, 'cover')
            
            # 解压源文档
            with zipfile.ZipFile(source_path, 'r') as zip_ref:
                zip_ref.extractall(source_temp)
            
            # 分析document.xml找到分割点
            doc_xml_path = os.path.join(source_temp, 'word', 'document.xml')
            if not os.path.exists(doc_xml_path):
                raise FileNotFoundError("未找到document.xml文件")
            
            # 读取并分析document.xml
            with open(doc_xml_path, 'r', encoding='utf-8') as f:
                doc_content = f.read()
            
            # 找到目录开始位置（使用封面文档的分割逻辑）
            split_point = self._find_split_point_cover(doc_content, toc_keywords)
            
            if split_point == -1:
                print("警告: 未找到明确的目录开始位置，使用默认分割点")
                split_point = self._get_default_split_point(doc_content)
            
            # 创建封面目录XML内容
            cover_xml = self._split_document_xml_cover(doc_content, split_point)
            
            # 复制完整的文档结构
            shutil.copytree(source_temp, cover_temp)
            
            # 写入分割后的document.xml
            with open(os.path.join(cover_temp, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
                f.write(cover_xml)
            
            # 重新打包为DOCX文件
            self._create_docx_from_temp(cover_temp, cover_toc_output)
            
            return True
            
        except Exception as e:
            print(f"处理封面目录文档时出错: {str(e)}")
            return False
        finally:
            # 清理临时文件
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
    
    def _process_content_no_toc_document(self, source_path: str, content_no_toc_output: str, toc_keywords: List[str]) -> bool:
        """
        处理生成不含目录的正文内容文档
        使用与docx_split_no_toc.py中相同的逻辑
        """
        temp_dir = None
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()
            source_temp = os.path.join(temp_dir, 'source')
            content_temp = os.path.join(temp_dir, 'content')
            
            # 解压源文档
            with zipfile.ZipFile(source_path, 'r') as zip_ref:
                zip_ref.extractall(source_temp)
            
            # 分析document.xml找到分割点
            doc_xml_path = os.path.join(source_temp, 'word', 'document.xml')
            if not os.path.exists(doc_xml_path):
                raise FileNotFoundError("未找到document.xml文件")
            
            # 读取并分析document.xml
            with open(doc_xml_path, 'r', encoding='utf-8') as f:
                doc_content = f.read()
            
            # 找到目录结束位置（使用不含目录文档的分割逻辑）
            split_point = self._find_split_point_content_no_toc(doc_content, toc_keywords)
            
            if split_point == -1:
                print("警告: 未找到明确的目录结束位置，使用默认分割点")
                split_point = self._get_default_split_point(doc_content)
            
            # 创建不含目录的正文内容XML
            content_xml = self._split_document_xml_content_no_toc(doc_content, split_point)
            
            # 复制完整的文档结构
            shutil.copytree(source_temp, content_temp)
            
            # 写入不含目录的正文内容document.xml
            with open(os.path.join(content_temp, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
                f.write(content_xml)
            
            # 重新打包为DOCX文件
            self._create_docx_from_temp(content_temp, content_no_toc_output)
            
            return True
            
        except Exception as e:
            print(f"处理正文文档时出错: {str(e)}")
            return False
        finally:
            # 清理临时文件
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
    
    def _find_split_point_cover(self, doc_content: str, toc_keywords: List[str]) -> int:
        """
        在XML内容中找到封面文档的分割点（目录开始位置）
        使用与docx_split_merge.py中相同的逻辑
        """
        try:
            # 找到所有文档主体元素（段落、表格、分节符等）
            element_patterns = [
                r'<w:p[^>]*>.*?</w:p>',          # 段落
                r'<w:tbl[^>]*>.*?</w:tbl>',       # 表格
                r'<w:sectPr[^>]*>.*?</w:sectPr>', # 分节符
                r'<w:bookmarkStart[^>]*/>',       # 书签开始
                r'<w:bookmarkEnd[^>]*/>',         # 书签结束
            ]
            
            # 合并所有元素模式
            combined_pattern = '|'.join(f'({pattern})' for pattern in element_patterns)
            
            # 找到所有文档元素及其位置
            elements = []
            for match in re.finditer(combined_pattern, doc_content, re.DOTALL):
                element_content = match.group(0)
                start_pos = match.start()
                end_pos = match.end()
                
                # 判断元素类型
                if element_content.startswith('<w:p'):
                    element_type = 'paragraph'
                elif element_content.startswith('<w:tbl'):
                    element_type = 'table'
                elif element_content.startswith('<w:sectPr'):
                    element_type = 'section'
                else:
                    element_type = 'other'
                
                elements.append({
                    'content': element_content,
                    'type': element_type,
                    'start': start_pos,
                    'end': end_pos,
                    'index': len(elements)
                })
            
            print(f"封面文档处理 - 找到 {len(elements)} 个文档元素")
            
            toc_start = -1
            
            # 查找目录开始位置（只在段落中查找）
            for i, element in enumerate(elements):
                if element['type'] == 'paragraph':
                    # 提取段落文本
                    text_pattern = r'<w:t[^>]*>([^<]*)</w:t>'
                    texts = re.findall(text_pattern, element['content'])
                    paragraph_text = ''.join(texts).strip()
                    
                    if any(keyword in paragraph_text for keyword in toc_keywords):
                        toc_start = i
                        print(f"封面文档处理 - 找到目录开始位置：元素 {i}，内容：{paragraph_text[:50]}...")
                        break
            
            if toc_start == -1:
                print("封面文档处理 - 未找到目录关键词")
                return -1
            
            # 对于封面部分，我们只需要目录开始之前的内容
            # 所以分割点就是目录开始的位置
            split_point = toc_start
            
            print(f"封面文档处理 - 设置分割点为目录开始位置：元素 {split_point}")
            
            return split_point
            
        except Exception as e:
            print(f"分析封面文档XML时出错: {e}")
            import traceback
            traceback.print_exc()
            return -1
    
    def _find_split_point_content_no_toc(self, doc_content: str, toc_keywords: List[str]) -> int:
        """
        在XML内容中找到不含目录正文文档的分割点（目录结束位置）
        使用与docx_split_no_toc.py中相同的逻辑
        """
        try:
            # 找到所有文档主体元素（段落、表格、分节符等）
            element_patterns = [
                r'<w:p[^>]*>.*?</w:p>',          # 段落
                r'<w:tbl[^>]*>.*?</w:tbl>',       # 表格
                r'<w:sectPr[^>]*>.*?</w:sectPr>', # 分节符
                r'<w:bookmarkStart[^>]*/>',       # 书签开始
                r'<w:bookmarkEnd[^>]*/>',         # 书签结束
            ]
            
            # 合并所有元素模式
            combined_pattern = '|'.join(f'({pattern})' for pattern in element_patterns)
            
            # 找到所有文档元素及其位置
            elements = []
            for match in re.finditer(combined_pattern, doc_content, re.DOTALL):
                element_content = match.group(0)
                start_pos = match.start()
                end_pos = match.end()
                
                # 判断元素类型
                if element_content.startswith('<w:p'):
                    element_type = 'paragraph'
                elif element_content.startswith('<w:tbl'):
                    element_type = 'table'
                elif element_content.startswith('<w:sectPr'):
                    element_type = 'section'
                else:
                    element_type = 'other'
                
                elements.append({
                    'content': element_content,
                    'type': element_type,
                    'start': start_pos,
                    'end': end_pos,
                    'index': len(elements)
                })
            
            print(f"正文文档处理 - 找到 {len(elements)} 个文档元素")
            
            toc_start = -1
            toc_end = -1
            
            # 查找目录开始位置（只在段落中查找）
            for i, element in enumerate(elements):
                if element['type'] == 'paragraph':
                    # 提取段落文本
                    text_pattern = r'<w:t[^>]*>([^<]*)</w:t>'
                    texts = re.findall(text_pattern, element['content'])
                    paragraph_text = ''.join(texts).strip()
                    
                    if any(keyword in paragraph_text for keyword in toc_keywords):
                        toc_start = i
                        print(f"正文文档处理 - 找到目录开始位置：元素 {i}，内容：{paragraph_text[:50]}...")
                        break
            
            if toc_start == -1:
                print("正文文档处理 - 未找到目录关键词")
                return -1
            
            # 查找目录结束位置（考虑表格边界）
            consecutive_numbered_lines = 0
            for i in range(toc_start + 1, len(elements)):
                element = elements[i]
                
                if element['type'] == 'paragraph':
                    text_pattern = r'<w:t[^>]*>([^<]*)</w:t>'
                    texts = re.findall(text_pattern, element['content'])
                    paragraph_text = ''.join(texts).strip()
                    
                    if self._has_page_number_pattern(paragraph_text):
                        consecutive_numbered_lines += 1
                    elif paragraph_text == "":
                        continue
                    elif consecutive_numbered_lines >= 3:
                        toc_end = i - 1
                        print(f"正文文档处理 - 找到目录结束位置：元素 {toc_end}")
                        break
                    elif len(paragraph_text) > 100 and not self._has_page_number_pattern(paragraph_text):
                        # 发现长文本段落，很可能是正文开始
                        toc_end = i - 1
                        print(f"正文文档处理 - 根据长文本判断目录结束位置：元素 {toc_end}")
                        break
                    else:
                        consecutive_numbered_lines = 0
                
                elif element['type'] == 'table':
                    # 遇到表格，检查是否在目录区域
                    if consecutive_numbered_lines > 0:
                        # 如果之前有目录条目，表格可能是目录的一部分
                        continue
                    else:
                        # 表格很可能是正文内容，结束目录区域
                        toc_end = i - 1
                        print(f"正文文档处理 - 遇到表格，目录结束位置：元素 {toc_end}")
                        break
            
            # 如果没有找到明确的结束位置，使用启发式方法
            if toc_end == -1:
                # 限制搜索范围，避免将整个文档都当作目录
                search_limit = min(toc_start + 50, len(elements))
                for i in range(toc_start + 1, search_limit):
                    element = elements[i]
                    if element['type'] == 'paragraph':
                        text_pattern = r'<w:t[^>]*>([^<]*)</w:t>'
                        texts = re.findall(text_pattern, element['content'])
                        paragraph_text = ''.join(texts).strip()
                        
                        if len(paragraph_text) > 100 and not self._has_page_number_pattern(paragraph_text):
                            toc_end = i - 1
                            print(f"正文文档处理 - 启发式判断目录结束位置：元素 {toc_end}")
                            break
                
                # 如果仍然没找到，使用默认值
                if toc_end == -1:
                    toc_end = min(toc_start + 20, len(elements) - 1)
                    print(f"正文文档处理 - 使用默认目录结束位置：元素 {toc_end}")
            
            return toc_end
            
        except Exception as e:
            print(f"分析不含目录正文文档XML时出错: {e}")
            import traceback
            traceback.print_exc()
            return -1
    
    def _get_default_split_point(self, doc_content: str) -> int:
        """
        获取默认分割点（前20%的文档元素）
        改进算法：考虑所有文档元素，不仅仅是段落
        
        Args:
            doc_content: document.xml的内容
            
        Returns:
            int: 默认分割点
        """
        try:
            # 找到所有文档主体元素
            element_patterns = [
                r'<w:p[^>]*>.*?</w:p>',          # 段落
                r'<w:tbl[^>]*>.*?</w:tbl>',       # 表格
                r'<w:sectPr[^>]*>.*?</w:sectPr>', # 分节符
                r'<w:bookmarkStart[^>]*/>',       # 书签开始
                r'<w:bookmarkEnd[^>]*/>',         # 书签结束
            ]
            
            combined_pattern = '|'.join(f'({pattern})' for pattern in element_patterns)
            elements = list(re.finditer(combined_pattern, doc_content, re.DOTALL))
            
            total_elements = len(elements)
            # 返回前20个元素或前20%的元素
            default_point = min(20, total_elements // 5)
            
            print(f"使用默认分割点：{default_point} / {total_elements} 个元素")
            return default_point
            
        except Exception as e:
            print(f"获取默认分割点时出错: {e}")
            return 10
    
    def _split_document_xml_cover(self, doc_content: str, split_point: int) -> str:
        """
        将document.xml按照分割点分为封面目录部分
        使用与docx_split_merge.py中相同的逻辑
        """
        try:
            # 找到所有文档主体元素
            element_patterns = [
                r'<w:p[^>]*>.*?</w:p>',          # 段落
                r'<w:tbl[^>]*>.*?</w:tbl>',       # 表格
                r'<w:sectPr[^>]*>.*?</w:sectPr>', # 分节符
                r'<w:bookmarkStart[^>]*/>',       # 书签开始
                r'<w:bookmarkEnd[^>]*/>',         # 书签结束
            ]
            
            combined_pattern = '|'.join(f'({pattern})' for pattern in element_patterns)
            
            # 找到所有元素及其位置
            elements = []
            last_end = 0
            
            for match in re.finditer(combined_pattern, doc_content, re.DOTALL):
                element_content = match.group(0)
                start_pos = match.start()
                end_pos = match.end()
                
                elements.append({
                    'content': element_content,
                    'start': start_pos,
                    'end': end_pos,
                    'index': len(elements)
                })
                last_end = end_pos
            
            if split_point < 0 or split_point >= len(elements):
                print(f"警告：分割点 {split_point} 超出范围，使用默认分割")
                split_point = min(len(elements) // 4, 20)  # 使用前1/4或前20个元素
            
            # 找到分割位置在原始文档中的位置
            if split_point < len(elements):
                split_pos = elements[split_point]['start']  # 使用start位置，确保不包含目录开始的段落
            else:
                split_pos = elements[-1]['end'] if elements else len(doc_content) // 2
            
            print(f"封面文档处理 - 在位置 {split_pos} 处分割文档（元素索引 {split_point}/{len(elements)}）")
            
            # 找到body标签的开始和结束位置
            body_start_pattern = r'<w:body[^>]*>'
            body_end_pattern = r'</w:body>'
            
            body_start_match = re.search(body_start_pattern, doc_content)
            body_end_match = re.search(body_end_pattern, doc_content)
            
            if not body_start_match or not body_end_match:
                print("警告：未找到body标签，使用简单分割")
                mid_point = len(doc_content) // 2
                return doc_content[:mid_point] + '</w:body></w:document>'
            
            body_start_end = body_start_match.end()
            body_end_start = body_end_match.start()
            
            # 确保分割点在body内部
            if split_pos < body_start_end:
                split_pos = body_start_end
                print(f"调整分割点到body开始位置: {split_pos}")
            elif split_pos > body_end_start:
                split_pos = body_end_start
                print(f"调整分割点到body结束位置: {split_pos}")
            
            # 获取XML框架部分
            xml_header = doc_content[:body_start_end]  # 从文档开始到body开始
            xml_footer = doc_content[body_end_start:]   # 从body结束到文档结束
            
            # 获取body内的内容
            body_content = doc_content[body_start_end:body_end_start]
            
            # 在body内容中进行分割
            body_split_pos = split_pos - body_start_end
            
            # 确保不会超出范围
            if body_split_pos < 0:
                body_split_pos = 0
            elif body_split_pos > len(body_content):
                body_split_pos = len(body_content)
            
            # 封面部分只包含目录之前的内容
            cover_body_content = body_content[:body_split_pos]
            
            # 构建封面目录XML（只包含目录之前的内容）
            cover_xml = xml_header + cover_body_content + xml_footer
            
            # 验证XML的完整性
            self._validate_xml_structure(cover_xml, "封面目录")
            
            print(f"封面文档处理 - 分割完成：封面目录 {len(cover_xml)} 字符")
            
            return cover_xml
            
        except Exception as e:
            print(f"分割封面XML内容时出错: {e}")
            import traceback
            traceback.print_exc()
            # 返回原始内容作为备用
            return doc_content[:len(doc_content)//2] + '</w:body></w:document>'
    
    def _split_document_xml_content_no_toc(self, doc_content: str, split_point: int) -> str:
        """
        将document.xml按照分割点分为不含目录的正文内容
        使用与docx_split_no_toc.py中相同的逻辑
        """
        try:
            # 找到所有文档主体元素
            element_patterns = [
                r'<w:p[^>]*>.*?</w:p>',          # 段落
                r'<w:tbl[^>]*>.*?</w:tbl>',       # 表格
                r'<w:sectPr[^>]*>.*?</w:sectPr>', # 分节符
                r'<w:bookmarkStart[^>]*/>',       # 书签开始
                r'<w:bookmarkEnd[^>]*/>',         # 书签结束
            ]
            
            combined_pattern = '|'.join(f'({pattern})' for pattern in element_patterns)
            
            # 找到所有元素及其位置
            elements = []
            last_end = 0
            
            for match in re.finditer(combined_pattern, doc_content, re.DOTALL):
                element_content = match.group(0)
                start_pos = match.start()
                end_pos = match.end()
                
                elements.append({
                    'content': element_content,
                    'start': start_pos,
                    'end': end_pos,
                    'index': len(elements)
                })
                last_end = end_pos
            
            if split_point < 0 or split_point >= len(elements):
                print(f"警告：分割点 {split_point} 超出范围，使用默认分割")
                split_point = min(len(elements) // 4, 20)  # 使用前1/4或前20个元素
            
            # 找到分割位置在原始文档中的位置
            # 对于正文部分，我们需要目录结束位置之后的内容（不包括目录本身）
            if split_point < len(elements):
                split_pos = elements[split_point]['end']  # 使用end位置，确保包含完整的目录结束位置
            else:
                split_pos = elements[-1]['end'] if elements else len(doc_content) // 2
            
            print(f"正文文档处理 - 在位置 {split_pos} 处分割文档（元素索引 {split_point}/{len(elements)}）")
            
            # 找到body标签的开始和结束位置
            body_start_pattern = r'<w:body[^>]*>'
            body_end_pattern = r'</w:body>'
            
            body_start_match = re.search(body_start_pattern, doc_content)
            body_end_match = re.search(body_end_pattern, doc_content)
            
            if not body_start_match or not body_end_match:
                print("警告：未找到body标签，使用简单分割")
                mid_point = len(doc_content) // 2
                return doc_content[:doc_content.find('<w:body')] + '<w:body>' + doc_content[mid_point:]
            
            body_start_end = body_start_match.end()
            body_end_start = body_end_match.start()
            
            # 确保分割点在body内部
            if split_pos < body_start_end:
                split_pos = body_start_end
                print(f"调整分割点到body开始位置: {split_pos}")
            elif split_pos > body_end_start:
                split_pos = body_end_start
                print(f"调整分割点到body结束位置: {split_pos}")
            
            # 获取XML框架部分
            xml_header = doc_content[:body_start_end]  # 从文档开始到body开始
            xml_footer = doc_content[body_end_start:]   # 从body结束到文档结束
            
            # 获取body内的内容
            body_content = doc_content[body_start_end:body_end_start]
            
            # 在body内容中进行分割
            body_split_pos = split_pos - body_start_end
            
            # 确保不会超出范围
            if body_split_pos < 0:
                body_split_pos = 0
            elif body_split_pos > len(body_content):
                body_split_pos = len(body_content)
            
            # 正文部分从目录结束位置之后开始（不包含目录）
            content_body_content = body_content[body_split_pos:]
            
            # 构建不含目录的正文内容XML（从目录结束位置之后开始，不包含目录）
            content_xml = xml_header + content_body_content + xml_footer
            
            # 验证XML的完整性
            self._validate_xml_structure(content_xml, "正文内容")
            
            print(f"正文文档处理 - 分割完成：不含目录的正文内容 {len(content_xml)} 字符")
            
            return content_xml
            
        except Exception as e:
            print(f"分割不含目录正文XML内容时出错: {e}")
            import traceback
            traceback.print_exc()
            # 返回原始内容作为备用
            return doc_content[len(doc_content)//2:]
    
    def _validate_xml_structure(self, xml_content: str, description: str):
        """
        验证XML结构的完整性
        
        Args:
            xml_content: XML内容
            description: 描述信息
        """
        try:
            import xml.etree.ElementTree as ET
            try:
                ET.fromstring(xml_content.encode('utf-8'))
                print(f"✅ {description} XML结构验证通过")
            except ET.ParseError as e:
                print(f"⚠️ {description} XML结构验证失败: {e}")
            except Exception as e:
                print(f"⚠️ {description} XML验证过程中出错: {e}")
        except Exception as e:
            print(f"⚠️ {description} XML验证初始化失败: {e}")
    
    def _create_docx_from_temp(self, temp_dir: str, output_path: str):
        """
        从临时目录创建DOCX文件
        
        Args:
            temp_dir: 临时目录路径
            output_path: 输出文件路径
        """
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 按照DOCX标准创建ZIP文件
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zip_file:
            # 按照DOCX标准顺序添加文件
            
            # 1. [Content_Types].xml
            content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
            if os.path.exists(content_types_path):
                zip_file.write(content_types_path, '[Content_Types].xml')
            
            # 2. _rels 目录
            rels_dir = os.path.join(temp_dir, '_rels')
            if os.path.exists(rels_dir):
                for root, dirs, files in os.walk(rels_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, temp_dir)
                        zip_file.write(file_path, arc_name)
            
            # 3. 所有其他文件
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    
                    # 跳过已经添加的文件
                    if arc_name != '[Content_Types].xml' and not arc_name.startswith('_rels'):
                        zip_file.write(file_path, arc_name)
    
    def _has_page_number_pattern(self, text: str) -> bool:
        """
        检查文本是否包含页码模式
        
        Args:
            text: 待检查文本
            
        Returns:
            bool: 是否包含页码模式
        """
        if not text:
            return False
        
        # 常见的页码模式
        patterns = [
            r'\.{2,}\s*\d+$',  # ...123
            r'\s+\d+$',        # 空格加数字结尾
            r'\d+$',           # 纯数字结尾
            r'…+\s*\d+$',      # 省略号加数字
        ]
        
        for pattern in patterns:
            if re.search(pattern, text):
                return True
        
        return False


def split_document_for_conversion(source_path: str,
                                output_dir: Optional[str] = None,
                                toc_keywords: Optional[List[str]] = None) -> Tuple[Optional[str], Optional[str]]:
    """
    便捷函数：将DOCX文档拆分为封面目录文档和不含目录的正文内容文档
    
    Args:
        source_path: 源文档路径
        output_dir: 输出目录（可选，默认为源文件所在目录）
        toc_keywords: 目录识别关键词列表（可选）
        
    Returns:
        Tuple[str, str]: (封面目录文件路径, 不含目录的正文内容文件路径)
        
    说明:
        - 封面目录文档：只包含目录之前的内容（如封面、标题等），不包含目录本身
        - 正文内容文档：包含从目录之后开始的所有内容（不包括目录）
    """
    if not os.path.exists(source_path):
        print(f"❌ 源文档不存在: {source_path}")
        return None, None
    
    if output_dir is None:
        output_dir = os.path.dirname(source_path)
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 获取文件名基础
    base_name = os.path.splitext(os.path.basename(source_path))[0]
    
    # 生成输出文件路径
    cover_toc_output = os.path.join(output_dir, f"{base_name}_封面目录.docx")
    content_no_toc_output = os.path.join(output_dir, f"{base_name}_正文内容.docx")
    
    # 创建处理器并执行拆分
    processor = DocxSplitProcessor()
    success = processor.split_document_for_conversion(
        source_path=source_path,
        cover_toc_output=cover_toc_output,
        content_no_toc_output=content_no_toc_output,
        toc_keywords=toc_keywords
    )
    
    if success:
        return cover_toc_output, content_no_toc_output
    else:
        return None, None


def quick_split_for_conversion(source_path: str,
                             output_dir: Optional[str] = None,
                             toc_keywords: Optional[List[str]] = None) -> Tuple[Optional[str], Optional[str]]:
    """
    便捷函数：快速将DOCX文档拆分为封面目录文档和不含目录的正文内容文档
    
    Args:
        source_path: 源文档路径
        output_dir: 输出目录（可选，默认为源文件所在目录）
        toc_keywords: 目录识别关键词列表（可选）
        
    Returns:
        Tuple[str, str]: (封面目录文件路径, 不含目录的正文内容文件路径)
        
    说明:
        - 封面目录文档：只包含目录之前的内容（如封面、标题等），不包含目录本身
        - 正文内容文档：包含从目录之后开始的所有内容（不包括目录）
    """
    return split_document_for_conversion(source_path, output_dir, toc_keywords)


# 保持向后兼容的别名
split_document = split_document_for_conversion
quick_split = quick_split_for_conversion


if __name__ == "__main__":
    # 示例用法
    print("DOCX文档拆分工具")
    print("=" * 40)
    
    # 测试拆分功能
    source_file = "test_document.docx"  # 请替换为实际文件路径
    
    if os.path.exists(source_file):
        cover_toc, content_no_toc = quick_split_for_conversion(source_file)
        
        if cover_toc and content_no_toc:
            print(f"拆分成功!")
            print(f"封面目录: {cover_toc}")
            print(f"正文内容(不含目录): {content_no_toc}")
        else:
            print("拆分失败!")
    else:
        print(f"测试文件 {source_file} 不存在，请提供一个真实的 DOCX 文件来测试此功能。")