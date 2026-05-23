import os
import sys
import requests
from pathlib import Path
import zipfile
from datetime import datetime
import json
from collections import defaultdict

# Windows平台Word转PDF
try:
    import win32com.client as win32
    WORD_AVAILABLE = True
except ImportError:
    WORD_AVAILABLE = False
    print("警告: win32com未安装，将尝试使用docx2pdf")
    try:
        from docx2pdf import convert
        DOCX2PDF_AVAILABLE = True
    except ImportError:
        DOCX2PDF_AVAILABLE = False
        print("错误: 请安装pywin32或docx2pdf库")

class WordToPDFParser:
    def __init__(self):
        self.base_dir = Path(__file__).parent
        self.pdf_dir = self.base_dir / "PDF_data"
        self.output_dir = self.base_dir / "parse_results"
        self.api_url = "http://xn-a.suanjiayun.com:55305/file_parse"
        
        # 创建必要的目录
        self.pdf_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
    
    def convert_word_to_pdf(self, word_file_path):
        """将Word文档转换为PDF"""
        word_path = Path(word_file_path)
        if not word_path.exists():
            raise FileNotFoundError(f"Word文件不存在: {word_file_path}")
        
        # 生成PDF文件路径
        pdf_filename = word_path.stem + ".pdf"
        pdf_path = self.pdf_dir / pdf_filename
        
        print(f"正在转换: {word_path.name} -> {pdf_filename}")
        
        if WORD_AVAILABLE and sys.platform == "win32":
            # 使用win32com转换（Windows平台最稳定）
            try:
                word = win32.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(str(word_path.absolute()))
                doc.SaveAs(str(pdf_path.absolute()), FileFormat=17)  # 17 = PDF格式
                doc.Close()
                word.Quit()
                print(f"✓ 转换成功: {pdf_filename}")
                return pdf_path
            except Exception as e:
                print(f"✗ win32com转换失败: {e}")
                if DOCX2PDF_AVAILABLE:
                    print("尝试使用docx2pdf...")
                else:
                    raise
        
        if DOCX2PDF_AVAILABLE:
            # 使用docx2pdf作为备选方案
            try:
                convert(str(word_path), str(pdf_path))
                print(f"✓ 转换成功: {pdf_filename}")
                return pdf_path
            except Exception as e:
                print(f"✗ docx2pdf转换失败: {e}")
                raise
        
        raise RuntimeError("无法转换Word文档，请安装pywin32或docx2pdf库")
    
    def parse_pdf(self, pdf_path, **kwargs):
        """调用API解析PDF文件"""
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
        
        print(f"\n正在解析PDF: {pdf_path.name}")
        
        # 准备表单数据（默认参数）
        form_data = {
            'output_dir': './output',
            'lang_list': 'ch',  # 中文
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
        
        # 更新自定义参数
        form_data.update(kwargs)
        
        # 准备文件
        with open(pdf_path, 'rb') as f:
            files = {'files': (pdf_path.name, f, 'application/pdf')}
            
            try:
                # 发送请求
                print("正在调用API...")
                response = requests.post(
                    self.api_url,
                    files=files,
                    data=form_data,
                    timeout=300  # 5分钟超时
                )
                
                if response.status_code == 200:
                    # 保存返回的ZIP文件
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    result_filename = f"{pdf_path.stem}_{timestamp}_result.zip"
                    result_path = self.output_dir / result_filename
                    
                    with open(result_path, 'wb') as f:
                        f.write(response.content)
                    
                    print(f"✓ 解析成功，结果已保存: {result_filename}")
                    
                    # 解压ZIP文件查看内容
                    extract_dir = self.output_dir / f"{pdf_path.stem}_{timestamp}"
                    extract_dir.mkdir(exist_ok=True)
                    
                    with zipfile.ZipFile(result_path, 'r') as zip_ref:
                        zip_ref.extractall(extract_dir)
                    print(f"✓ 结果已解压到: {extract_dir.name}")
                    
                    # 自动格式化content_list.json文件
                    self.format_content_list_in_dir(extract_dir)
                    
                    return result_path, extract_dir
                else:
                    print(f"✗ API调用失败: 状态码 {response.status_code}")
                    print(f"响应内容: {response.text}")
                    return None, None
                    
            except requests.exceptions.Timeout:
                print("✗ API调用超时")
                return None, None
            except Exception as e:
                print(f"✗ API调用出错: {e}")
                return None, None
    
    def format_content_list_in_dir(self, extract_dir):
        """格式化解压目录中的content_list.json文件"""
        extract_dir = Path(extract_dir)
        
        # 查找所有content_list.json文件
        json_files = list(extract_dir.rglob("*_content_list.json"))
        
        if not json_files:
            return
        
        print(f"\n正在格式化content_list.json文件...")
        
        for json_file in json_files:
            try:
                # 读取JSON文件
                with open(json_file, 'r', encoding='utf-8') as f:
                    content_list = json.load(f)
                
                # 按页码组织内容
                pages_content = defaultdict(list)
                
                for item in content_list:
                    page_idx = item.get('page_idx', 0)
                    item_type = item.get('type', '')
                    
                    # 跳过页码类型
                    if item_type == 'page_number':
                        continue
                    
                    # 提取需要的字段
                    content_item = {}
                    
                    # 表格标题
                    if 'table_caption' in item and item['table_caption']:
                        content_item['table_caption'] = item['table_caption']
                    
                    # 表格内容
                    if 'table_body' in item and item['table_body']:
                        content_item['table_body'] = item['table_body']
                    
                    # 图片路径
                    if 'img_path' in item and item['img_path']:
                        content_item['img_path'] = item['img_path']
                    
                    # 文本内容（如果有）
                    if 'text' in item and item['text']:
                        content_item['text'] = item['text']
                    
                    # 只添加有内容的项
                    if content_item:
                        content_item['type'] = item_type
                        pages_content[page_idx].append(content_item)
                
                # 生成格式化输出
                output_lines = []
                img_counter = {}  # 用于记录每页的图片计数
                
                for page_idx in sorted(pages_content.keys()):
                    page_num = page_idx + 1  # 页码从1开始
                    output_lines.append(f"<page {page_num}>")
                    img_counter[page_num] = 0  # 初始化当前页的图片计数
                    
                    for item in pages_content[page_idx]:
                        # 添加表格标题（只保留内容）
                        if 'table_caption' in item:
                            captions = item['table_caption']
                            if isinstance(captions, list):
                                for caption in captions:
                                    output_lines.append(caption)
                            else:
                                output_lines.append(captions)
                        
                        # 添加表格内容（只保留内容，前后添加空行）
                        if 'table_body' in item:
                            output_lines.append("")  # 表格前空行
                            output_lines.append(item['table_body'])
                            output_lines.append("")  # 表格后空行
                        
                        # 添加文本内容（只保留内容）
                        if 'text' in item:
                            output_lines.append(item['text'])
                        
                        # 添加图片路径（使用Markdown格式）
                        if 'img_path' in item:
                            img_counter[page_num] += 1
                            img_path = item['img_path']
                            # 生成新的图片名称：page{页码}_img{序号}
                            img_name = f"page{page_num}_img{img_counter[page_num]}"
                            # 获取原始图片扩展名
                            img_ext = Path(img_path).suffix or '.jpg'
                            # 生成Markdown格式的图片引用
                            output_lines.append(f"![{img_name}{img_ext}]({img_path})")
                        
                        # 项之间添加空行
                        output_lines.append("")
                    
                    output_lines.append(f"</page {page_num}>")
                    output_lines.append("")  # 页面之间添加空行
                
                # 生成输出文件路径（保存为Markdown格式）
                output_path = json_file.parent / f"{json_file.stem}_formatted.md"
                
                # 写入文件
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(output_lines))
                
                print(f"✓ 格式化完成: {json_file.name} -> {output_path.name}")
                print(f"  共 {len(pages_content)} 页")
                
            except Exception as e:
                print(f"✗ 格式化失败 {json_file.name}: {e}")
    
    def process_file(self, file_path):
        """处理文件：支持Word和PDF文件"""
        file_path = Path(file_path)
        print("="*50)
        print(f"开始处理: {file_path.name}")
        print("="*50)
        
        try:
            # 判断文件类型
            if file_path.suffix.lower() == '.pdf':
                # 如果是PDF，直接解析
                print(f"检测到PDF文件，直接进行解析...")
                pdf_path = file_path
            elif file_path.suffix.lower() in ['.doc', '.docx']:
                # 如果是Word，先转换再解析
                print(f"检测到Word文件，先进行转换...")
                pdf_path = self.convert_word_to_pdf(file_path)
            else:
                print(f"✗ 不支持的文件格式: {file_path.suffix}")
                return False
            
            # 解析PDF
            result_zip, extract_dir = self.parse_pdf(pdf_path)
            
            if result_zip:
                print("\n✓ 处理完成！")
                print(f"  - PDF文件: {pdf_path}")
                print(f"  - 结果ZIP: {result_zip}")
                print(f"  - 解压目录: {extract_dir}")
                return True
            else:
                print("\n✗ 处理失败")
                return False
                
        except Exception as e:
            print(f"\n✗ 处理出错: {e}")
            return False
    
    # 保留旧的方法名以保持兼容性
    def process_word_file(self, word_file_path):
        """兼容旧接口：处理Word文件"""
        return self.process_file(word_file_path)


def main():
    """主函数"""
    parser = WordToPDFParser()
    
    # 选择扫描位置
    print("选择文件扫描位置：")
    print("1. PDF转换+提取文件夹（当前目录）")
    print("2. PDF_data子文件夹（已转换的PDF）")
    print("3. 父目录（19-文件解析）")
    print("4. 全部位置")
    
    location_choice = input("请选择 (1/2/3/4，默认为4): ").strip() or "4"
    
    word_files = []
    pdf_files = []
    current_dir = Path(__file__).parent
    
    if location_choice in ["1", "4"]:
        # 扫描当前目录
        word_files.extend(current_dir.glob("*.docx"))
        word_files.extend(current_dir.glob("*.doc"))
        pdf_files.extend(current_dir.glob("*.pdf"))
    
    if location_choice in ["2", "4"]:
        # 扫描PDF_data子文件夹
        pdf_data_dir = current_dir / "PDF_data"
        if pdf_data_dir.exists():
            pdf_files.extend(pdf_data_dir.glob("*.pdf"))
    
    if location_choice in ["3", "4"]:
        # 扫描父目录
        parent_dir = current_dir.parent
        word_files.extend(parent_dir.glob("*.docx"))
        word_files.extend(parent_dir.glob("*.doc"))
        pdf_files.extend(parent_dir.glob("*.pdf"))
    
    # 去重
    word_files = list(set(word_files))
    pdf_files = list(set(pdf_files))
    all_files = word_files + pdf_files
    
    if not all_files:
        print("\n未找到Word或PDF文件")
        return
    
    print(f"找到 {len(all_files)} 个文件：")
    print("\n[Word文件]")
    for i, file in enumerate(word_files, 1):
        print(f"{i}. {file.name}")
    
    if pdf_files:
        print("\n[PDF文件]")
        for i, file in enumerate(pdf_files, len(word_files) + 1):
            print(f"{i}. {file.name}")
    
    print("\n选择处理方式：")
    print("1. 处理所有文件")
    print("2. 选择特定文件")
    print("3. 只处理Word文件")
    print("4. 只处理PDF文件")
    
    choice = input("请输入选择 (1/2/3/4): ").strip()
    
    if choice == "1":
        # 处理所有文件
        for file in all_files:
            parser.process_file(file)
            print("\n" + "="*50 + "\n")
    elif choice == "2":
        # 选择特定文件
        file_num = input(f"请输入文件编号 (1-{len(all_files)}): ").strip()
        try:
            idx = int(file_num) - 1
            if 0 <= idx < len(all_files):
                parser.process_file(all_files[idx])
            else:
                print("无效的文件编号")
        except ValueError:
            print("请输入有效的数字")
    elif choice == "3":
        # 只处理Word文件
        if word_files:
            for file in word_files:
                parser.process_file(file)
                print("\n" + "="*50 + "\n")
        else:
            print("没有找到Word文件")
    elif choice == "4":
        # 只处理PDF文件
        if pdf_files:
            for file in pdf_files:
                parser.process_file(file)
                print("\n" + "="*50 + "\n")
        else:
            print("没有找到PDF文件")
    else:
        print("无效的选择")


if __name__ == "__main__":
    main()
