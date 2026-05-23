#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档解析服务 - 双路径处理
路径1: MinerU解析PDF（不保存图片）
路径2: PDF转图片 + VL模型分析文字间逻辑关系
最终融合两条路径结果
"""
import os
import sys
import requests
import shutil
import zipfile
import json
import re
import base64
import concurrent.futures
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# Windows平台Word转PDF
WORD_AVAILABLE = False
DOCX2PDF_AVAILABLE = False
PDF2IMAGE_AVAILABLE = False
FITZ_AVAILABLE = False

try:
    import win32com.client as win32
    import pythoncom  # 用于多线程COM初始化
    WORD_AVAILABLE = True
except ImportError:
    pass

try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    pass

# PDF转图片库
try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    pass

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    pass

# VL模型调用
try:
    from langchain_openai import ChatOpenAI
    from langchain_core.messages import HumanMessage
    VL_AVAILABLE = True
except ImportError:
    VL_AVAILABLE = False


class DocumentParser:
    """文档解析器：双路径处理，支持docx/pdf解析"""
    
    def __init__(self, output_base_dir=None, vl_api_key=None, vl_base_url=None, vl_model=None):
        """
        初始化文档解析器
        
        Args:
            output_base_dir: 输出基础目录，默认为项目目录下的 parsed_documents
            vl_api_key: VL模型API密钥
            vl_base_url: VL模型API地址
            vl_model: VL模型名称
        """
        self.base_dir = Path(__file__).parent
        self.output_base_dir = Path(output_base_dir) if output_base_dir else self.base_dir / "parsed_documents"
        self.temp_dir = self.base_dir / "temp_parsing"
        self.api_url = "http://xn-a.suanjiayun.com:55305/file_parse"
        
        # VL模型配置
        self.vl_api_key = vl_api_key or os.getenv("QWEN_VL_API_KEY", "EMPTY")
        self.vl_base_url = vl_base_url or os.getenv("QWEN_VL_BASE_URL", "http://xn-d.suanjiayun.com:57915/v1")
        self.vl_model = vl_model or os.getenv("QWEN_VL_MODEL", "Qwen/Qwen3-VL-235B-A22B-Instruct-FP8")
        
        # 初始化VL模型客户端
        self.vl_client = None
        if VL_AVAILABLE:
            try:
                self.vl_client = ChatOpenAI(
                    model=self.vl_model,
                    openai_api_key=self.vl_api_key,
                    openai_api_base=self.vl_base_url,
                    temperature=0.1,
                    max_tokens=4000,
                    request_timeout=120
                )
                print(f"✅ VL模型已初始化: {self.vl_model}")
            except Exception as e:
                print(f"⚠️ VL模型初始化失败: {e}")
        
        # 创建必要的目录
        self.output_base_dir.mkdir(exist_ok=True)
        self.temp_dir.mkdir(exist_ok=True)
    
    def convert_docx_to_pdf(self, docx_path: Path) -> Path:
        """
        将docx转换为pdf
        
        Args:
            docx_path: docx文件路径
            
        Returns:
            pdf文件路径
        """
        docx_path = Path(docx_path)
        if not docx_path.exists():
            raise FileNotFoundError(f"文件不存在: {docx_path}")
        
        pdf_path = self.temp_dir / (docx_path.stem + ".pdf")
        
        print(f"📄 正在转换: {docx_path.name} -> {pdf_path.name}")
        
        if WORD_AVAILABLE and sys.platform == "win32":
            try:
                # 在多线程环境中需要初始化COM
                pythoncom.CoInitialize()
                try:
                    word = win32.Dispatch("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(str(docx_path.absolute()))
                    doc.SaveAs(str(pdf_path.absolute()), FileFormat=17)
                    doc.Close()
                    word.Quit()
                    print(f"✅ 转换成功: {pdf_path.name}")
                    return pdf_path
                finally:
                    pythoncom.CoUninitialize()
            except Exception as e:
                print(f"⚠️ win32com转换失败: {e}")
                if DOCX2PDF_AVAILABLE:
                    print("尝试使用docx2pdf...")
                else:
                    raise
        
        if DOCX2PDF_AVAILABLE:
            try:
                docx2pdf_convert(str(docx_path), str(pdf_path))
                print(f"✅ 转换成功: {pdf_path.name}")
                return pdf_path
            except Exception as e:
                print(f"❌ docx2pdf转换失败: {e}")
                raise
        
        raise RuntimeError("无法转换Word文档，请安装pywin32或docx2pdf库")
    
    def pdf_to_images(self, pdf_path: Path, output_dir: Path) -> list:
        """
        将PDF每一页保存为图片
        
        Args:
            pdf_path: PDF文件路径
            output_dir: 图片输出目录
            
        Returns:
            图片路径列表 [(page_num, image_path), ...]
        """
        pdf_path = Path(pdf_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(exist_ok=True)
        
        images = []
        
        print(f"📷 正在将PDF转换为图片...")
        
        if FITZ_AVAILABLE:
            # 使用PyMuPDF
            try:
                doc = fitz.open(str(pdf_path))
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    # 使用较高DPI获取清晰图片
                    mat = fitz.Matrix(2, 2)  # 2x缩放
                    pix = page.get_pixmap(matrix=mat)
                    
                    img_name = f"第{page_num + 1}页.png"
                    img_path = output_dir / img_name
                    pix.save(str(img_path))
                    
                    images.append((page_num + 1, img_path))
                    print(f"  ✅ 已保存: {img_name}")
                
                doc.close()
                print(f"✅ PDF转图片完成，共 {len(images)} 页")
                return images
            except Exception as e:
                print(f"⚠️ PyMuPDF转换失败: {e}")
        
        if PDF2IMAGE_AVAILABLE:
            try:
                pil_images = convert_from_path(str(pdf_path), dpi=200)
                for page_num, pil_img in enumerate(pil_images, 1):
                    img_name = f"第{page_num}页.png"
                    img_path = output_dir / img_name
                    pil_img.save(str(img_path), 'PNG')
                    images.append((page_num, img_path))
                    print(f"  ✅ 已保存: {img_name}")
                
                print(f"✅ PDF转图片完成，共 {len(images)} 页")
                return images
            except Exception as e:
                print(f"⚠️ pdf2image转换失败: {e}")
        
        print("❌ 无法将PDF转换为图片，请安装PyMuPDF或pdf2image库")
        return []
    
    def _encode_image(self, image_path: Path) -> str:
        """将图片编码为base64"""
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode('utf-8')
    
    def analyze_page_with_vl(self, page_num: int, image_path: Path) -> dict:
        """
        使用VL模型分析单页图片，识别文字间的逻辑关系
        
        Args:
            page_num: 页码
            image_path: 图片路径
            
        Returns:
            {
                'page_num': int,
                'has_special_layout': bool,  # 是否有特殊布局（如板书设计）
                'layout_type': str,  # 布局类型
                'layout_description': str,  # 布局描述
                'text_relationships': str,  # 文字间关系描述
                'raw_response': str  # 原始响应
            }
        """
        if not self.vl_client:
            return {
                'page_num': page_num,
                'has_special_layout': False,
                'layout_type': '',
                'layout_description': '',
                'text_relationships': '',
                'raw_response': 'VL模型未初始化'
            }
        
        try:
            base64_image = self._encode_image(image_path)
            
            prompt = """请分析这张教案文档图片，重点关注以下内容：

1. 识别是否存在"板书设计"、"教学流程图"、"思维导图"、"结构图"等具有文字间逻辑关系的特殊布局
2. 如果存在这类特殊布局，请详细描述：
   - 布局类型（如：板书设计、流程图、树状图等）
   - 文字内容及其位置关系（上下左右、层级、箭头指向等）
   - 文字之间的逻辑关系（如：包含关系、因果关系、并列关系、递进关系等）
   - 用结构化的文字描述来还原这个布局的完整含义

3. 如果只是普通文本段落或表格，简单说明"无特殊布局"即可

请按以下JSON格式返回：
{
    "has_special_layout": true/false,
    "layout_type": "布局类型，如：板书设计、流程图、无",
    "layout_description": "布局的整体描述",
    "text_relationships": "详细的文字间关系描述，用文字还原布局的结构和逻辑"
}

请严格按JSON格式返回，不要添加其他内容。"""
            
            message = HumanMessage(
                content=[
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{base64_image}"}
                    }
                ]
            )
            
            response = self.vl_client.invoke([message])
            raw_response = response.content
            
            # 解析JSON响应
            try:
                # 提取JSON部分
                json_match = re.search(r'\{[\s\S]*\}', raw_response)
                if json_match:
                    result = json.loads(json_match.group())
                    return {
                        'page_num': page_num,
                        'has_special_layout': result.get('has_special_layout', False),
                        'layout_type': result.get('layout_type', ''),
                        'layout_description': result.get('layout_description', ''),
                        'text_relationships': result.get('text_relationships', ''),
                        'raw_response': raw_response
                    }
            except json.JSONDecodeError:
                pass
            
            # 如果JSON解析失败，返回原始响应
            return {
                'page_num': page_num,
                'has_special_layout': '板书' in raw_response or '流程' in raw_response or '结构' in raw_response,
                'layout_type': '待解析',
                'layout_description': '',
                'text_relationships': raw_response,
                'raw_response': raw_response
            }
            
        except Exception as e:
            print(f"  ⚠️ 第{page_num}页VL分析失败: {e}")
            return {
                'page_num': page_num,
                'has_special_layout': False,
                'layout_type': '',
                'layout_description': '',
                'text_relationships': '',
                'raw_response': str(e)
            }
    
    # 修改 analyze_all_pages_with_vl 方法：

    def analyze_all_pages_with_vl(self, page_images, progress_callback=None) -> list:
        """
        使用VL模型分析所有页面（兼容列表和字典格式）
        
        Args:
            page_images: 可以是列表 [(page_num, image_path), ...] 
                        也可以是字典 {page_num: (page_num, image_path), ...}
            progress_callback: 进度回调
            
        Returns:
            分析结果列表
        """
        # 统一转换为字典格式
        page_images_dict = {}
        
        if isinstance(page_images, list):
            # 列表格式：[(page_num, image_path), ...]
            print(f"📋 检测到列表格式，共 {len(page_images)} 项")
            for i, item in enumerate(page_images):
                if isinstance(item, tuple) and len(item) == 2:
                    page_num, image_path = item
                    page_images_dict[page_num] = (page_num, image_path)
                elif hasattr(item, 'page_num') and hasattr(item, 'image_path'):
                    # 如果是对象
                    page_images_dict[item.page_num] = (item.page_num, item.image_path)
        elif isinstance(page_images, dict):
            # 字典格式：{page_num: (page_num, image_path), ...}
            print(f"📋 检测到字典格式，共 {len(page_images)} 项")
            page_images_dict = page_images
        else:
            print(f"❌ 不支持的page_images格式: {type(page_images)}")
            return []
        
        results = []
        total = len(page_images_dict)
        
        if total == 0:
            print("⚠️ 没有图片需要分析")
            return results
        
        print(f"🤖 开始VL模型分析，共 {total} 页...")
        
        # 按页码排序
        sorted_pages = sorted(page_images_dict.items())
        
        for idx, (page_num, (_, image_path)) in enumerate(sorted_pages):
            print(f"  分析第 {page_num:03d} 页 ({idx + 1}/{total})...", end='', flush=True)
            
            if progress_callback:
                # 计算进度：30% 开始，30% 用于分析
                progress = 30 + int((idx / total) * 30)
                progress_callback(progress, f"VL分析第{page_num}页...")
            
            result = self.analyze_page_with_vl(page_num, image_path)
            results.append(result)
            
            if result['has_special_layout']:
                print(f" ✅ 发现特殊布局: {result['layout_type']}")
            else:
                print(f" ✅ 无特殊布局")
        
        # 统计特殊布局
        special_count = sum(1 for r in results if r['has_special_layout'])
        print(f"✅ VL分析完成，发现 {special_count} 个特殊布局")
        
        return results
    
    def parse_document_per_page(self, file_path: str, progress_callback=None, enable_vl_analysis: bool = True) -> dict:
        """
        逐页解析文档 - 确保分页准确
        
        Args:
            file_path: 文件路径（docx或pdf）
            progress_callback: 进度回调函数
            enable_vl_analysis: 是否启用VL模型分析
            
        Returns:
            解析结果
        """
        file_path = Path(file_path)
        doc_name = file_path.stem
        
        print("=" * 60)
        print(f"📄 开始逐页解析文档: {file_path.name}")
        print("=" * 60)
        
        if progress_callback:
            progress_callback(5, f"开始逐页处理: {file_path.name}")
        
        try:
            suffix = file_path.suffix.lower()
            
            # 创建输出目录
            output_dir = self.output_base_dir / doc_name
            output_dir.mkdir(exist_ok=True)
            
            # 单页PDF目录
            single_pdf_dir = output_dir / "single_pdfs"
            single_pdf_dir.mkdir(exist_ok=True)
            
            # 页面图片目录
            page_images_dir = output_dir / "page_images"
            page_images_dir.mkdir(exist_ok=True)
            
            if suffix in ['.doc', '.docx']:
                # Word文档：逐页转换为PDF
                if progress_callback:
                    progress_callback(10, "正在逐页转换Word文档...")
                
                # 转换并拆分为单页PDF
                single_pdf_paths = self.convert_docx_to_pdf_per_page(file_path, single_pdf_dir)
                
                if not single_pdf_paths:
                    return {
                        'success': False,
                        'error': "Word转换PDF失败"
                    }
                    
            elif suffix == '.pdf':
                # PDF文件：直接拆分为单页
                if progress_callback:
                    progress_callback(10, "正在拆分为单页PDF...")
                
                if not FITZ_AVAILABLE:
                    return {
                        'success': False,
                        'error': "需要PyMuPDF库来拆分PDF"
                    }
                
                # 拆分PDF为单页
                single_pdf_paths = self._split_pdf_to_single_pages(file_path, single_pdf_dir)
                
            else:
                return {
                    'success': False,
                    'error': f"不支持的文件格式: {suffix}"
                }
            
            print(f"\n📊 共生成 {len(single_pdf_paths)} 个单页PDF")
            
            # ========== 路径2: 逐页PDF转图片 + VL分析 ==========
            vl_results = []
            page_images_dict = {}
            
            if enable_vl_analysis and self.vl_client:
                print("\n" + "-" * 40)
                print("🔀 路径2: 逐页PDF转图片 + VL模型分析")
                print("-" * 40)
                
                if progress_callback:
                    progress_callback(20, "正在将单页PDF转换为图片...")
                
                # 将每个单页PDF转换为图片
                page_images_dict = self.pdf_to_images_per_page(single_pdf_paths, page_images_dir)
                
                if page_images_dict:
                    # VL模型分析每页
                    if progress_callback:
                        progress_callback(30, "正在进行VL模型分析...")
                    
                    vl_results = self.analyze_all_pages_with_vl(page_images_dict, progress_callback)
            else:
                print("⚠️ VL模型未启用或未初始化，跳过路径2")
            
            # ========== 路径1: 逐页调用MinerU解析 ==========
            print("\n" + "-" * 40)
            print("🔀 路径1: 逐页调用MinerU解析")
            print("-" * 40)
            
            all_pages_content = {}
            
            for idx, pdf_path in enumerate(single_pdf_paths):
                page_num = idx + 1
                
                if progress_callback:
                    progress_callback(60 + int((idx / len(single_pdf_paths)) * 15), 
                                    f"正在解析第{page_num}页...")
                
                print(f"\n📄 解析第 {page_num:03d}/{len(single_pdf_paths)} 页...")
                
                # 调用MinerU API解析单页PDF
                zip_path, extract_dir = self.call_mineru_api(pdf_path)
                
                if not extract_dir:
                    print(f"  ❌ 第{page_num}页解析失败")
                    continue
                
                # 处理解析结果
                _, _, _, page_content = self.process_mineru_result(
                    extract_dir, f"{doc_name}_page{page_num:03d}", save_images=False
                )
                
                if page_content:
                    # page_content的键是0（因为单页PDF只有一页）
                    if 0 in page_content:
                        all_pages_content[page_num - 1] = page_content[0]  # 转换为从0开始的索引
                        print(f"  ✅ 第{page_num}页解析成功")
            
            # ========== 生成融合的Markdown文件 ==========
            if progress_callback:
                progress_callback(85, "正在生成Markdown文件...")
            
            md_path = self._generate_merged_markdown(doc_name, output_dir, all_pages_content)
            
            if not md_path:
                return {
                    'success': False,
                    'error': "生成Markdown文件失败"
                }
            
            # ========== 融合VL分析结果 ==========
            if progress_callback:
                progress_callback(90, "正在融合VL分析结果...")
            
            # 将VL分析结果融合到Markdown
            if vl_results:
                self.merge_vl_analysis(md_path, vl_results)
            
            # 统计
            page_image_count = len(page_images_dict)
            
            if progress_callback:
                progress_callback(100, "解析完成!")
            
            print("\n" + "=" * 60)
            print("✅ 逐页解析完成!")
            print(f"📁 输出目录: {output_dir}")
            print(f"📝 Markdown文件: {md_path}")
            print(f"📄 单页PDF数量: {len(single_pdf_paths)}")
            print(f"🖼️  页面图片数量: {page_image_count}")
            print(f"🤖 VL分析特殊布局: {sum(1 for r in vl_results if r.get('has_special_layout'))}")
            print("=" * 60)
            
            return {
                'success': True,
                'doc_name': doc_name,
                'output_dir': str(output_dir),
                'md_path': str(md_path),
                'single_pdf_dir': str(single_pdf_dir),
                'page_images_dir': str(page_images_dir),
                'image_count': page_image_count,
                'vl_analysis': vl_results,
                'total_pages': len(single_pdf_paths)
            }
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'error': str(e)
            }

    def _generate_merged_markdown(self, doc_name: str, output_dir: Path, all_pages_content: dict) -> Path:
        """
        生成融合所有页面内容的Markdown文件
        
        Args:
            doc_name: 文档名称
            output_dir: 输出目录
            all_pages_content: 所有页面的内容字典
            
        Returns:
            Markdown文件路径
        """
        md_path = output_dir / f"{doc_name}.md"
        
        output_lines = []
        
        # 按页码排序
        sorted_pages = sorted(all_pages_content.items())
        
        for page_idx, page_items in sorted_pages:
            page_num = page_idx + 1
            output_lines.append(f"<page {page_num}>")
            
            img_counter = 0
            
            for item in page_items:
                if 'table_caption' in item:
                    captions = item['table_caption']
                    if isinstance(captions, list):
                        for caption in captions:
                            output_lines.append(caption)
                    else:
                        output_lines.append(captions)
                
                if 'table_body' in item:
                    output_lines.append("")
                    output_lines.append(item['table_body'])
                    output_lines.append("")
                
                if 'text' in item:
                    output_lines.append(item['text'])
                
                if 'img_path' in item:
                    img_counter += 1
                    # 生成新的图片名称
                    img_ext = Path(item['img_path']).suffix or '.png'
                    new_img_name = f"image_{page_num}_{img_counter}{img_ext}"
                    #output_lines.append(f"[图片: {new_img_name}]")
                
                output_lines.append("")
            
            output_lines.append(f"</page {page_num}>")
            output_lines.append("")
        
        # 写入md文件
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(output_lines))
        
        print(f"✅ 生成Markdown: {md_path.name}")
        print(f"   共 {len(all_pages_content)} 页")
        
        return md_path

    def call_mineru_api(self, pdf_path: Path, save_images: bool = True) -> tuple:
        """
        调用MinerU API解析PDF
        
        Args:
            pdf_path: PDF文件路径
            
        Returns:
            (zip_path, extract_dir) 或 (None, None) 失败时
        """
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
        
        print(f"🔍 正在调用MinerU API解析: {pdf_path.name}")
        
        # 准备表单数据
        form_data = {
            'output_dir': './output',
            'lang_list': 'ch',
            'backend': 'vlm-vllm-async-engine',
            'parse_method': 'txt',
            'formula_enable': 'true',
            'table_enable': 'true',
            'server_url': 'string',
            'return_md': 'true',
            'return_middle_json': 'true',
            'return_model_output': 'true',
            'return_content_list': 'true',
            'return_images': 'true',
            'response_format_zip': 'true',
            'start_page_id': '0',
            'end_page_id': '99999'
        }
        
        with open(pdf_path, 'rb') as f:
            files = {'files': (pdf_path.name, f, 'application/pdf')}
            
            try:
                print("⏳ 正在解析（可能需要几分钟）...")
                response = requests.post(
                    self.api_url,
                    files=files,
                    data=form_data,
                    timeout=600  # 10分钟超时
                )
                
                if response.status_code == 200:
                    # 保存ZIP文件
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    zip_filename = f"{pdf_path.stem}_{timestamp}_result.zip"
                    zip_path = self.temp_dir / zip_filename
                    
                    with open(zip_path, 'wb') as zf:
                        zf.write(response.content)
                    
                    print(f"✅ API解析成功")
                    
                    # 解压ZIP
                    extract_dir = self.temp_dir / f"{pdf_path.stem}_{timestamp}"
                    extract_dir.mkdir(exist_ok=True)
                    
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        zip_ref.extractall(extract_dir)
                    
                    print(f"✅ 解压完成: {extract_dir.name}")
                    return zip_path, extract_dir
                else:
                    print(f"❌ API调用失败: 状态码 {response.status_code}")
                    print(f"响应: {response.text[:500]}")
                    return None, None
                    
            except requests.exceptions.Timeout:
                print("❌ API调用超时")
                return None, None
            except Exception as e:
                print(f"❌ API调用出错: {e}")
                return None, None
    # 在 DocumentParser 类中添加以下方法：

    def convert_docx_to_pdf_per_page(self, docx_path: Path, output_dir: Path = None) -> list[Path]:
        """
        将docx逐页转换为PDF（每页一个PDF文件）
        
        Args:
            docx_path: docx文件路径
            output_dir: 输出目录，默认使用temp_dir
            
        Returns:
            单页PDF文件路径列表，按页码顺序
        """
        docx_path = Path(docx_path)
        if not docx_path.exists():
            raise FileNotFoundError(f"文件不存在: {docx_path}")
        
        if output_dir is None:
            output_dir = self.temp_dir / f"{docx_path.stem}_single_pages"
        
        output_dir = Path(output_dir)
        output_dir.mkdir(exist_ok=True)
        
        print(f"📄 正在逐页转换: {docx_path.name}")
        
        single_pdf_paths = []
        
        # 回退方案：先转换为完整PDF，然后拆分为单页
        return self._split_pdf_to_single_pages(docx_path, output_dir)


    def _split_pdf_to_single_pages(self, docx_path: Path, output_dir: Path) -> list[Path]:
        """
        将Word转换为PDF，然后拆分为单页PDF
        
        步骤：
        1. 将Word转换为完整的PDF
        2. 使用PyMuPDF将完整PDF拆分为每页一个的PDF文件
        
        Args:
            docx_path: Word文件路径
            output_dir: 输出目录
            
        Returns:
            单页PDF文件路径列表
        """
        print(f"  使用回退方案：先转换完整PDF再拆分...")
        
        # 1. 先转换为完整PDF
        full_pdf_path = self.convert_docx_to_pdf(docx_path)
        
        if not FITZ_AVAILABLE:
            print(f"  ❌ 未安装PyMuPDF，无法拆分PDF")
            return [full_pdf_path]  # 返回完整PDF作为备选
        
        # 2. 使用PyMuPDF拆分PDF
        single_pdf_paths = []
        
        try:
            doc = fitz.open(str(full_pdf_path))
            total_pages = len(doc)
            print(f"  📊 PDF总页数: {total_pages}")
            
            for page_num in range(total_pages):
                # 创建只包含当前页的新PDF
                single_pdf = fitz.open()
                single_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
                
                # 保存单页PDF
                single_pdf_path = output_dir / f"page_{page_num + 1:03d}.pdf"
                single_pdf.save(str(single_pdf_path))
                single_pdf.close()
                
                single_pdf_paths.append(single_pdf_path)
                print(f"  ✅ 已保存第 {page_num + 1:03d}/{total_pages} 页: {single_pdf_path.name}")
            
            doc.close()
            
            # 清理完整PDF
            if full_pdf_path.exists():
                try:
                    full_pdf_path.unlink()
                    print(f"  🗑️ 已清理完整PDF: {full_pdf_path.name}")
                except:
                    pass
            
            print(f"  🎉 拆分完成，共 {len(single_pdf_paths)} 个单页PDF")
            return single_pdf_paths
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"  ❌ PDF拆分失败: {e}")
            return [full_pdf_path]  # 返回完整PDF作为备选


    def pdf_to_images_per_page(self, pdf_paths: list[Path], output_dir: Path) -> dict[int, tuple[int, Path]]:
        """
        将多个单页PDF转换为图片（每页对应一张图片）
        
        Args:
            pdf_paths: 单页PDF文件路径列表
            output_dir: 图片输出目录
            
        Returns:
            字典: {页码: (原始页码, 图片路径)}
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(exist_ok=True)
        
        page_images = {}
        
        print(f"📷 正在将 {len(pdf_paths)} 个单页PDF转换为图片...")
        
        if not FITZ_AVAILABLE:
            print(f"❌ 未安装PyMuPDF，无法转换PDF为图片")
            return page_images
        
        for pdf_path in pdf_paths:
            try:
                # 从文件名提取页码
                filename = pdf_path.stem
                page_num_match = re.search(r'page_(\d+)', filename)
                if page_num_match:
                    page_num = int(page_num_match.group(1))
                else:
                    page_num = len(page_images) + 1
                
                # 使用PyMuPDF打开单页PDF
                doc = fitz.open(str(pdf_path))
                
                if len(doc) == 1:  # 确保是单页PDF
                    page = doc.load_page(0)
                    # 使用较高DPI获取清晰图片
                    mat = fitz.Matrix(2, 2)  # 2x缩放
                    pix = page.get_pixmap(matrix=mat)
                    
                    # 生成图片文件名
                    img_name = f"第{page_num:03d}页.png"
                    img_path = output_dir / img_name
                    pix.save(str(img_path))
                    
                    page_images[page_num] = (page_num, img_path)
                    print(f"  ✅ 第{page_num:03d}页: {img_name}")
                else:
                    print(f"  ⚠️  {pdf_path.name} 不是单页PDF，跳过")
                
                doc.close()
                
            except Exception as e:
                print(f"  ❌ 转换失败 {pdf_path.name}: {e}")
                continue
        
        print(f"✅ PDF转图片完成，共 {len(page_images)} 页")
        return page_images

    def process_mineru_result(self, extract_dir: Path, doc_name: str, save_images: bool = False) -> tuple:
        """
        处理MinerU解析结果，生成标准格式的.md文件
        
        Args:
            extract_dir: 解压后的目录
            doc_name: 文档名称（不含扩展名）
            save_images: 是否保存图片（默认False，双路径模式不需要保存）
            
        Returns:
            (output_dir, md_path, images_dir, pages_content) 输出目录、md文件路径、图片目录、页面内容
        """
        extract_dir = Path(extract_dir)
        
        # 创建输出目录
        output_dir = self.output_base_dir / doc_name
        output_dir.mkdir(exist_ok=True)
        
        md_path = output_dir / f"{doc_name}.md"
        images_dir = output_dir / "images"
        if save_images:
            images_dir.mkdir(exist_ok=True)
        
        print(f"📁 输出目录: {output_dir}")
        
        # 查找content_list.json文件
        json_files = list(extract_dir.rglob("*_content_list.json"))
        
        if not json_files:
            # 如果没有content_list.json，查找.md文件
            md_files = list(extract_dir.rglob("*.md"))
            if md_files:
                # 直接复制md文件
                shutil.copy(md_files[0], md_path)
                print(f"⚠️ 未找到content_list.json，直接复制.md文件")
            else:
                print(f"❌ 未找到解析结果")
                return None, None, None
        else:
            # 处理content_list.json
            json_file = json_files[0]
            print(f"📝 处理: {json_file.name}")
            
            with open(json_file, 'r', encoding='utf-8') as f:
                content_list = json.load(f)
            
            # 按页码组织内容
            pages_content = defaultdict(list)
            img_mapping = {}  # 原始图片路径 -> 新图片路径
            
            for item in content_list:
                page_idx = item.get('page_idx', 0)
                item_type = item.get('type', '')
                
                if item_type == 'page_number':
                    continue
                
                content_item = {}
                
                if 'table_caption' in item and item['table_caption']:
                    content_item['table_caption'] = item['table_caption']
                
                if 'table_body' in item and item['table_body']:
                    content_item['table_body'] = item['table_body']
                
                if 'text' in item and item['text']:
                    content_item['text'] = item['text']
                
                if 'img_path' in item and item['img_path']:
                    content_item['img_path'] = item['img_path']
                
                if content_item:
                    content_item['type'] = item_type
                    pages_content[page_idx].append(content_item)
            
            # 生成Markdown内容
            output_lines = []
            img_counter = {}
            
            for page_idx in sorted(pages_content.keys()):
                page_num = page_idx + 1
                output_lines.append(f"<page {page_num}>")
                img_counter[page_num] = 0
                
                for item in pages_content[page_idx]:
                    if 'table_caption' in item:
                        captions = item['table_caption']
                        if isinstance(captions, list):
                            for caption in captions:
                                output_lines.append(caption)
                        else:
                            output_lines.append(captions)
                    
                    if 'table_body' in item:
                        output_lines.append("")
                        output_lines.append(item['table_body'])
                        output_lines.append("")
                    
                    if 'text' in item:
                        output_lines.append(item['text'])
                    
                    if 'img_path' in item:
                        img_counter[page_num] += 1
                        original_img_path = item['img_path']
                        
                        # 生成新的图片名称
                        img_ext = Path(original_img_path).suffix or '.png'
                        new_img_name = f"image_{page_num}_{img_counter[page_num]}{img_ext}"
                        
                        if save_images:
                            new_img_path = images_dir / new_img_name
                            
                            # 复制图片到images目录
                            original_full_path = extract_dir / original_img_path
                            if not original_full_path.exists():
                                # 尝试在子目录中查找
                                found_imgs = list(extract_dir.rglob(Path(original_img_path).name))
                                if found_imgs:
                                    original_full_path = found_imgs[0]
                            
                            if original_full_path.exists():
                                shutil.copy(original_full_path, new_img_path)
                                img_mapping[original_img_path] = new_img_name
                            
                            # 使用相对路径引用图片
                            output_lines.append(f"![{new_img_name}](images/{new_img_name})")
                        else:
                            # 不保存图片，只添加占位符
                            output_lines.append(f"[图片: {new_img_name}]")
                    
                    output_lines.append("")
                
                output_lines.append(f"</page {page_num}>")
                output_lines.append("")
            
            # 写入md文件
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(output_lines))
            
            print(f"✅ 生成Markdown: {md_path.name}")
            print(f"   共 {len(pages_content)} 页，{sum(img_counter.values())} 张图片")
        
        # 复制其他图片（如果需要且目录存在）
        if save_images and images_dir.exists():
            for img_file in extract_dir.rglob("*.png"):
                if img_file.name not in [p.name for p in images_dir.iterdir()]:
                    shutil.copy(img_file, images_dir / img_file.name)
            for img_file in extract_dir.rglob("*.jpg"):
                if img_file.name not in [p.name for p in images_dir.iterdir()]:
                    shutil.copy(img_file, images_dir / img_file.name)
            for img_file in extract_dir.rglob("*.jpeg"):
                if img_file.name not in [p.name for p in images_dir.iterdir()]:
                    shutil.copy(img_file, images_dir / img_file.name)
        
        return output_dir, md_path, images_dir, dict(pages_content)
    
    def merge_vl_analysis(self, md_path: Path, vl_results: list) -> Path:
        """
        将VL模型分析结果融合到Markdown文档中
        
        Args:
            md_path: MinerU生成的Markdown文件路径
            vl_results: VL模型分析结果列表
            
        Returns:
            融合后的Markdown文件路径
        """
        md_path = Path(md_path)
        
        # 读取原始Markdown内容
        with open(md_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
        
        # 筛选有特殊布局的页面
        special_layouts = [r for r in vl_results if r.get('has_special_layout')]
        
        if not special_layouts:
            print("📝 没有发现需要融合的特殊布局")
            return md_path
        
        print(f"🔄 开始融合 {len(special_layouts)} 个特殊布局分析...")
        
        # 对每个有特殊布局的页面进行融合
        for layout in special_layouts:
            page_num = layout['page_num']
            layout_type = layout.get('layout_type', '特殊布局')
            text_relationships = layout.get('text_relationships', '')
            
            if not text_relationships:
                continue
            
            # 构建VL分析补充内容
            vl_supplement = f"""

<!-- VL模型分析补充 - 第{page_num}页 -->
**📊 {layout_type}结构分析：**

{text_relationships}

<!-- 补充结束 -->
"""
            
            # 在对应页面的结束标签前插入补充内容
            page_end_tag = f"</page {page_num}>"
            if page_end_tag in md_content:
                md_content = md_content.replace(
                    page_end_tag,
                    vl_supplement + page_end_tag
                )
                print(f"  ✅ 已融合第{page_num}页: {layout_type}")
        
        # 写入融合后的内容
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(md_content)
        
        print(f"✅ 融合完成，已更新: {md_path.name}")
        return md_path
    
    def parse_document(self, file_path: str, progress_callback=None, 
                    enable_vl_analysis: bool = True, use_per_page: bool = False) -> dict:
        """
        解析文档（主入口） - 支持两种模式
        
        Args:
            file_path: 文件路径
            progress_callback: 进度回调
            enable_vl_analysis: 是否启用VL模型分析
            use_per_page: 是否使用逐页转换模式（默认False，保持兼容性）
            
        Returns:
            解析结果
        """
        try:
            if use_per_page:
                # 调用逐页解析方法
                result = self.parse_document_per_page(
                    file_path,
                    progress_callback=progress_callback,
                    enable_vl_analysis=enable_vl_analysis
                )
            else:
                # 调用标准解析方法
                result = self._parse_document_original(
                    file_path,
                    progress_callback=progress_callback,
                    enable_vl_analysis=enable_vl_analysis
                )
            
            # ✅ 解析完成后自动清理临时文件
            print("\n🧹 解析完成，清理临时文件...")
            self.cleanup_temp()  # 完全清理
            
            return result
            
        except Exception as e:
            # ❌ 解析失败时也尝试清理
            print("🧹 解析失败，清理临时文件...")
            self.cleanup_temp()
            raise e

    def _parse_document_original(self, file_path: str, progress_callback=None, enable_vl_analysis: bool = True) -> dict:
        """
        解析文档（主入口） - 双路径处理
        
        路径1: MinerU解析PDF（不保存图片）
        路径2: PDF转图片 + VL模型分析文字间逻辑关系
        最终融合两条路径结果
        
        Args:
            file_path: 文件路径（docx或pdf）
            progress_callback: 进度回调函数 callback(percent, message)
            enable_vl_analysis: 是否启用VL模型分析（默认True）
            
        Returns:
            {
                'success': bool,
                'doc_name': str,
                'output_dir': str,
                'md_path': str,
                'images_dir': str,
                'page_images_dir': str,  # 页面图片目录
                'image_count': int,
                'vl_analysis': list,  # VL分析结果
                'error': str (如果失败)
            }
        """
        file_path = Path(file_path)
        doc_name = file_path.stem
        
        print("=" * 60)
        print(f"📄 开始双路径解析文档: {file_path.name}")
        print("=" * 60)
        
        if progress_callback:
            progress_callback(5, f"开始处理: {file_path.name}")
        
        try:
            # 1. 判断文件类型，获取PDF
            suffix = file_path.suffix.lower()
            
            if suffix in ['.doc', '.docx']:
                # Word文档，先转换为PDF
                if progress_callback:
                    progress_callback(8, "正在转换Word文档为PDF...")
                pdf_path = self.convert_docx_to_pdf(file_path)
            elif suffix == '.pdf':
                pdf_path = file_path
                if progress_callback:
                    progress_callback(8, "检测到PDF文件...")
            else:
                return {
                    'success': False,
                    'error': f"不支持的文件格式: {suffix}"
                }
            
            # 创建输出目录
            output_dir = self.output_base_dir / doc_name
            output_dir.mkdir(exist_ok=True)
            page_images_dir = output_dir / "page_images"
            page_images_dir.mkdir(exist_ok=True)
            
            # ========== 路径2: PDF转图片 + VL分析 ==========
            vl_results = []
            page_images = []
            
            if enable_vl_analysis and self.vl_client:
                print("\n" + "-" * 40)
                print("🔀 路径2: PDF转图片 + VL模型分析")
                print("-" * 40)
                
                if progress_callback:
                    progress_callback(10, "正在将PDF转换为图片...")
                
                # 2.1 PDF转图片
                page_images = self.pdf_to_images(pdf_path, page_images_dir)
                
                if page_images:
                    # 2.2 VL模型分析每页
                    if progress_callback:
                        progress_callback(20, "正在进行VL模型分析...")
                    
                    vl_results = self.analyze_all_pages_with_vl(page_images, progress_callback)
            else:
                print("⚠️ VL模型未启用或未初始化，跳过路径2")
            
            # ========== 路径1: MinerU解析 ==========
            print("\n" + "-" * 40)
            print("🔀 路径1: MinerU模型解析")
            print("-" * 40)
            
            if progress_callback:
                progress_callback(60, "正在调用MinerU模型解析...")
            
            # 3. 调用MinerU API解析（不保存图片）
            zip_path, extract_dir = self.call_mineru_api(pdf_path)
            
            if not extract_dir:
                return {
                    'success': False,
                    'error': "MinerU API解析失败"
                }
            
            # 4. 处理解析结果（不保存图片）
            if progress_callback:
                progress_callback(75, "正在生成Markdown文件...")
            
            result = self.process_mineru_result(extract_dir, doc_name, save_images=False)
            output_dir, md_path, images_dir, pages_content = result
            
            if not md_path:
                return {
                    'success': False,
                    'error': "处理解析结果失败"
                }
            
            # ========== 融合两条路径结果 ==========
            print("\n" + "-" * 40)
            print("🔄 融合双路径结果")
            print("-" * 40)
            
            if progress_callback:
                progress_callback(85, "正在融合VL分析结果...")
            
            # 5. 将VL分析结果融合到Markdown
            if vl_results:
                self.merge_vl_analysis(md_path, vl_results)
            
            # 6. 清理临时文件
            if progress_callback:
                progress_callback(95, "正在清理临时文件...")
            
            # 统计
            page_image_count = len(page_images)
            
            if progress_callback:
                progress_callback(100, "解析完成!")
            
            print("\n" + "=" * 60)
            print("✅ 双路径文档解析完成!")
            print(f"📁 输出目录: {output_dir}")
            print(f"📝 融合后Markdown: {md_path}")
            print(f"🖼️  页面图片数量: {page_image_count}")
            print(f"🤖 VL分析特殊布局: {sum(1 for r in vl_results if r.get('has_special_layout'))}")
            print("=" * 60)
            
            return {
                'success': True,
                'doc_name': doc_name,
                'output_dir': str(output_dir),
                'md_path': str(md_path),
                'images_dir': str(images_dir) if images_dir else '',
                'page_images_dir': str(page_images_dir),
                'image_count': page_image_count,
                'vl_analysis': vl_results
            }
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'error': str(e)
            }
    
    def cleanup_temp(self):
        """清理临时文件"""
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
            self.temp_dir.mkdir(exist_ok=True)
            print("🗑️  已清理临时文件")


# 在 document_parser.py 文件末尾替换测试代码为：

# 辅助函数：打印目录树
def print_directory_tree(path: Path, level: int = 0, prefix: str = ""):
    """
    打印目录树结构
    
    Args:
        path: 目录路径
        level: 当前层级
        prefix: 前缀字符串
    """
    if not path.exists():
        return
    
    # 当前目录/文件
    indent = "│   " * (level - 1) + "├── " if level > 0 else ""
    print(prefix + indent + path.name)
    
    if path.is_dir():
        # 获取子项并按类型排序
        items = sorted(path.iterdir(), key=lambda x: (not x.is_dir(), x.name.lower()))
        for i, item in enumerate(items):
            is_last = i == len(items) - 1
            next_prefix = prefix + ("    " if level == 0 else ("│   " if not is_last else "    "))
            print_directory_tree(item, level + 1, next_prefix)


# 测试代码 - 指定输出目录
if __name__ == "__main__":
    # 指定输出目录
    custom_output_dir = r"./parsed_documents"
    
    # 确保目录存在
    os.makedirs(custom_output_dir, exist_ok=True)
    
    print("=" * 60)
    print(f"📁 输出目录: {custom_output_dir}")
    print("=" * 60)
    
    # 创建解析器，指定输出目录
    parser = DocumentParser(output_base_dir=custom_output_dir)
    
    # 清理临时文件（可选）
    parser.cleanup_temp()
    
    # 测试文件路径
    test_file = input("请输入要解析的文件路径 (支持 .docx, .doc, .pdf): ").strip()
    
    if not test_file:
        # 如果没有输入，尝试在脚本目录查找测试文件
        script_dir = Path(__file__).parent
        test_files = list(script_dir.glob("*.docx")) + list(script_dir.glob("*.pdf"))
        
        if test_files:
            print("\n找到以下文件:")
            for i, f in enumerate(test_files, 1):
                print(f"{i}. {f.name}")
            
            choice = input(f"\n请选择文件编号 (1-{len(test_files)}): ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(test_files):
                test_file = str(test_files[int(choice) - 1])
            else:
                print("❌ 无效选择")
                sys.exit(1)
        else:
            print("❌ 未找到测试文件")
            sys.exit(1)
    
    test_path = Path(test_file)
    if not test_path.exists():
        print(f"❌ 文件不存在: {test_file}")
        sys.exit(1)
    
    # 选择解析模式
    print("\n请选择解析模式:")
    print("1. 标准模式 (整个文档转换)")
    print("2. 逐页模式 (逐页转换，分页更准确)")
    
    mode_choice = input("请选择模式 (1/2，默认2): ").strip()
    if not mode_choice:
        mode_choice = "2"
    
    use_per_page = (mode_choice == "2")
    
    # 选择是否启用VL分析
    vl_choice = input("是否启用VL模型分析? (y/n，默认y): ").strip().lower()
    if not vl_choice:
        vl_choice = "y"
    enable_vl = (vl_choice != "n")
    
    print("\n" + "=" * 60)
    print("🚀 开始解析...")
    print(f"📄 文件: {test_path.name}")
    print(f"🔧 模式: {'逐页模式' if use_per_page else '标准模式'}")
    print(f"🤖 VL分析: {'启用' if enable_vl else '禁用'}")
    print("=" * 60)
    
    # 解析进度回调函数
    def progress_callback(percent, message):
        print(f"📊 [{percent:3d}%] {message}")
    
    try:
        # 开始解析
        if use_per_page:
            # 调用逐页解析方法
            result = parser.parse_document_per_page(
                str(test_path),
                progress_callback=progress_callback,
                enable_vl_analysis=enable_vl
            )
        else:
            # 调用标准解析方法
            result = parser.parse_document(
                str(test_path),
                progress_callback=progress_callback,
                enable_vl_analysis=enable_vl,
                use_per_page=False
            )
        
        print("\n" + "=" * 60)
        print("📊 解析结果:")
        print("=" * 60)
        
        if result['success']:
            print(f"✅ 解析成功!")
            print(f"📁 输出目录: {result['output_dir']}")
            print(f"📄 Markdown文件: {result['md_path']}")
            
            if 'total_pages' in result:
                print(f"📑 总页数: {result['total_pages']}")
            
            print(f"🖼️  图片数量: {result['image_count']}")
            
            if result.get('vl_analysis'):
                special_layouts = sum(1 for r in result['vl_analysis'] if r.get('has_special_layout'))
                print(f"🤖 VL分析特殊布局: {special_layouts}")
            
            # 显示文件结构
            print(f"\n📂 生成的文件结构:")
            output_path = Path(result['output_dir'])
            print_directory_tree(output_path)
            
            # 询问是否打开输出目录
            open_choice = input("\n是否打开输出目录? (y/n): ").strip().lower()
            if open_choice == 'y':
                try:
                    import subprocess
                    subprocess.Popen(f'explorer "{output_path}"')
                except:
                    print("⚠️ 无法打开目录")
            
        else:
            print(f"❌ 解析失败: {result.get('error', '未知错误')}")
            
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"❌ 解析过程出错: {e}")
    
    finally:
        # 清理临时文件
        parser.cleanup_temp()
