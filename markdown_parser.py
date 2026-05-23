"""
Markdown文档解析器
用于提取markdown文档中的段落、图片及其上下文
"""
import re
import os
from typing import List, Dict, Tuple
from dataclasses import dataclass


@dataclass
class TextSegment:
    """文本段落"""
    content: str
    line_start: int
    line_end: int
    paragraph_index: int
    page_number: int = 1  # 页码
    segment_type: str = "text"


@dataclass
class ImageSegment:
    """图片段落"""
    image_path: str
    alt_text: str
    line_number: int
    paragraph_index: int
    context_before: str  # 图片前的上下文
    context_after: str   # 图片后的上下文
    page_number: int = 1  # 页码
    segment_type: str = "image"


class MarkdownParser:
    """Markdown解析器"""
    
    def __init__(self, markdown_file_path: str, image_folder: str = "image"):
        """
        初始化解析器
        
        Args:
            markdown_file_path: markdown文件路径
            image_folder: 图片文件夹名称
        """
        self.markdown_file_path = markdown_file_path
        self.image_folder = image_folder
        self.base_dir = os.path.dirname(markdown_file_path)
        
    def parse(self) -> Tuple[List[TextSegment], List[ImageSegment]]:
        """
        解析markdown文档（支持<page N>标签）
        
        Returns:
            (文本段落列表, 图片段落列表)
        """
        with open(self.markdown_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            lines = content.split('\n')
        
        text_segments = []
        image_segments = []
        
        # 提取页码边界
        page_boundaries = self._extract_page_boundaries(lines)
        
        # 按段落分割（空行分隔）
        paragraphs = self._split_into_paragraphs(content)
        
        current_line = 0
        for para_idx, paragraph in enumerate(paragraphs):
            para_lines = paragraph.split('\n')
            para_line_count = len(para_lines)
            
            # 确定当前段落的页码
            current_page = self._get_page_number(current_line, page_boundaries)
            
            # 检查段落中是否包含图片
            images_in_para = self._extract_images(paragraph)
            
            if images_in_para:
                # 处理包含图片的段落
                for img_info in images_in_para:
                    # 获取图片的上下文
                    context_before, context_after = self._get_image_context(
                        paragraph, img_info['markdown']
                    )
                    
                    # 构建完整的图片路径
                    img_relative_path = img_info['path']
                    
                    # 方案1：markdown中的路径相对于markdown文档
                    img_path_relative_to_doc = os.path.join(self.base_dir, img_relative_path)
                    
                    # 方案3：使用--image-folder参数 + 只使用文件名
                    img_filename = os.path.basename(img_relative_path)
                    img_path_filename_only = os.path.join(self.base_dir, self.image_folder, img_filename)
                    
                    # 按顺序尝试，找到第一个存在的
                    if os.path.exists(img_path_relative_to_doc):
                        img_path = img_path_relative_to_doc
                    elif os.path.exists(img_path_filename_only):
                        img_path = img_path_filename_only
                    else:
                        # 都不存在，使用默认的方案1
                        img_path = img_path_relative_to_doc
                    
                    image_seg = ImageSegment(
                        image_path=img_path,
                        alt_text=img_info['alt'],
                        line_number=current_line + img_info['line_offset'],
                        paragraph_index=para_idx,
                        context_before=context_before,
                        context_after=context_after,
                        page_number=current_page
                    )
                    image_segments.append(image_seg)
            
            # 提取纯文本内容（移除markdown图片语法和页码标签）
            text_content = self._remove_image_syntax(paragraph)
            text_content = self._remove_page_tags(text_content).strip()
            
            if text_content:
                text_seg = TextSegment(
                    content=text_content,
                    line_start=current_line,
                    line_end=current_line + para_line_count - 1,
                    paragraph_index=para_idx,
                    page_number=current_page
                )
                text_segments.append(text_seg)
            
            current_line += para_line_count + 1  # +1 for empty line separator
        
        return text_segments, image_segments
    
    def _extract_page_boundaries(self, lines: List[str]) -> Dict[int, int]:
        """
        提取页码边界（<page N>标签）
        
        Args:
            lines: 文档的所有行
            
        Returns:
            字典: {line_number: page_number}
        """
        page_map = {}
        current_page = 1
        
        # 匹配 <page N> 标签
        page_start_pattern = re.compile(r'<page\s+(\d+)>', re.IGNORECASE)
        
        for line_idx, line in enumerate(lines):
            match = page_start_pattern.search(line)
            if match:
                current_page = int(match.group(1))
            page_map[line_idx] = current_page
        
        return page_map
    
    def _get_page_number(self, line_number: int, page_boundaries: Dict[int, int]) -> int:
        """
        根据行号获取页码
        
        Args:
            line_number: 行号
            page_boundaries: 页码边界字典
            
        Returns:
            页码
        """
        if not page_boundaries:
            return 1
        
        # 找到最接近的页码
        for line_idx in range(line_number, -1, -1):
            if line_idx in page_boundaries:
                return page_boundaries[line_idx]
        
        return 1
    
    def _remove_page_tags(self, text: str) -> str:
        """移除页码标签"""
        # 移除 <page N> 和 </page N> 标签
        text = re.sub(r'</?page\s*\d*>', '', text, flags=re.IGNORECASE)
        return text
    
    def _split_into_paragraphs(self, content: str) -> List[str]:
        """将内容按段落分割"""
        # 预处理：检测并修复换行符加倍问题
        # 如果 \n\n 的数量异常多（超过总换行符的一半），说明换行符被加倍了
        newline_count = content.count('\n')
        double_newline_count = content.count('\n\n')
        
        if newline_count > 0 and double_newline_count / newline_count > 0.4:
            print(f"    [调试] 检测到换行符加倍 (\n: {newline_count}, \n\n: {double_newline_count})，正在修复...")
            # 将 \n\n 替换回 \n
            content = content.replace('\n\n', '\n')
            # 将 \n\n\n\n (原本的\n\n) 替换回 \n\n
            # 注意：上面的替换已经把 \n\n\n\n 变成了 \n\n，所以不需要额外操作
            # 但为了保险，我们重新统计一下
            print("[调试] 修复后:\n: " + str(content.count('\n')) + ",\n\n:" + str(content.count('\n\n')))
        
        # 使用两个或多个连续换行符作为段落分隔符（允许中间有空白字符）
        # 修改：使用更严格的匹配，至少要有一个完全的空行
        paragraphs = re.split(r'\n[ \t]*\n+', content)
        print(f"    [调试] split后: {len(paragraphs)} 个原始段落")
        
        # 过滤和合并段落
        filtered_paragraphs = []
        skipped_count = 0
        for p in paragraphs:
            p = p.strip()
            if not p:
                skipped_count += 1
                continue
            
            # 跳过只包含页码标签的段落
            if re.match(r'^</?page\s*\d*>$', p, re.IGNORECASE):
                skipped_count += 1
                continue
            
            # 移除段落中的页码标签
            p = self._remove_page_tags(p)
            p = p.strip()
            
            if p:
                filtered_paragraphs.append(p)
            else:
                skipped_count += 1
        
        print(f"    [调试] 过滤后: {len(filtered_paragraphs)} 个段落 (跳过 {skipped_count} 个)")
        return filtered_paragraphs
    
    def _extract_images(self, paragraph: str) -> List[Dict]:
        """
        从段落中提取图片信息
        
        Returns:
            包含图片信息的字典列表，每个字典包含：
            - markdown: 完整的markdown图片语法
            - alt: 图片alt文本
            - path: 图片路径
            - line_offset: 在段落中的行偏移
        """
        images = []
        # Markdown图片语法: ![alt text](image_path)
        pattern = r'!\[(.*?)\]\((.*?)\)'
        
        lines = paragraph.split('\n')
        for line_offset, line in enumerate(lines):
            matches = re.finditer(pattern, line)
            for match in matches:
                images.append({
                    'markdown': match.group(0),
                    'alt': match.group(1),
                    'path': match.group(2),
                    'line_offset': line_offset
                })
        
        return images
    
    def _get_image_context(self, paragraph: str, image_markdown: str, 
                          context_length: int = 200) -> Tuple[str, str]:
        """
        获取图片的上下文
        
        Args:
            paragraph: 包含图片的段落
            image_markdown: 图片的markdown语法
            context_length: 上下文的最大字符数
            
        Returns:
            (图片前的上下文, 图片后的上下文)
        """
        # 移除所有markdown图片语法以获取纯文本
        text_only = self._remove_image_syntax(paragraph)
        
        # 找到图片在原文中的位置
        img_pos = paragraph.find(image_markdown)
        
        if img_pos == -1:
            return "", ""
        
        # 获取图片前后的文本
        before_img = paragraph[:img_pos]
        after_img = paragraph[img_pos + len(image_markdown):]
        
        # 移除markdown语法
        before_text = self._remove_image_syntax(before_img).strip()
        after_text = self._remove_image_syntax(after_img).strip()
        
        # 限制长度
        if len(before_text) > context_length:
            before_text = "..." + before_text[-context_length:]
        if len(after_text) > context_length:
            after_text = after_text[:context_length] + "..."
        
        return before_text, after_text
    
    def _remove_image_syntax(self, text: str) -> str:
        """移除markdown图片语法"""
        # 移除图片语法但保留alt文本
        text = re.sub(r'!\[(.*?)\]\(.*?\)', r'\1', text)
        return text


def get_page_number_from_line(line_number: int, lines_per_page: int = 50) -> int:
    """
    根据行号估算页码
    
    Args:
        line_number: 行号（从0开始）
        lines_per_page: 每页的行数（默认50行）
        
    Returns:
        页码（从1开始）
    """
    return (line_number // lines_per_page) + 1
