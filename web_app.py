#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
教案合规检测 Web 应用
提供API接口和Web界面
"""
import os
import sys
import json
import time
import tempfile
import shutil
from pathlib import Path
from flask import Flask, render_template, request, jsonify, Response, stream_with_context
from flask_cors import CORS
from dotenv import load_dotenv

# 强制重新加载模块（避免缓存问题）
import importlib
if 'markdown_parser' in sys.modules:
    importlib.reload(sys.modules['markdown_parser'])

# 导入现有模块
from markdown_parser import MarkdownParser
from compliance_checker import (
    ComplianceRulesLoader,
    TextComplianceChecker,
    ImageComplianceChecker
)
from report_generator import ReportGenerator
from document_parser import DocumentParser

# 加载环境变量
load_dotenv()

app = Flask(__name__)
CORS(app)  # 允许跨域请求

# 全局变量存储检测器和解析器
text_checker = None
image_checker = None
doc_parser = None  # 文档解析器


def initialize_checkers():
    """初始化检测器"""
    global text_checker, image_checker
    
    # 加载API配置
    deepseek_api_key = os.getenv("DEEPSEEK_API_KEY")
    deepseek_base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
    qwen_vl_api_key = os.getenv("QWEN_VL_API_KEY", "EMPTY")
    qwen_vl_base_url = os.getenv("QWEN_VL_BASE_URL", "http://xn-d.suanjiayun.com:57915/v1")
    qwen_vl_model = os.getenv("QWEN_VL_MODEL", "Qwen/Qwen3-VL-235B-A22B-Instruct-FP8")
    
    if not deepseek_api_key:
        print("⚠️  警告: 未设置 DEEPSEEK_API_KEY，部分功能可能无法使用")
        return False
    
    # 加载合规规则
    rules = ComplianceRulesLoader("compliance_rules.json")
    
    # 初始化检测器
    text_checker = TextComplianceChecker(deepseek_api_key, deepseek_base_url, rules)
    image_checker = ImageComplianceChecker(qwen_vl_api_key, qwen_vl_base_url, rules, qwen_vl_model)
    
    return True


@app.route('/')
def index():
    """首页"""
    return render_template('index.html')


@app.route('/favicon.ico')
def favicon():
    """返回favicon，避免404错误"""
    return '', 204  # 204 No Content


@app.route('/api/test', methods=['GET'])
def test_connection():
    """测试API连接"""
    print("🔔 收到测试请求")
    return jsonify({
        'status': 'ok',
        'message': '服务器运行正常',
        'timestamp': __import__('datetime').datetime.now().isoformat()
    })


@app.route('/api/upload', methods=['POST'])
def upload_and_parse():
    """
    上传文档并解析（支持docx/pdf）
    双路径处理：
    - 路径1: MinerU解析（不保存图片）
    - 路径2: PDF转图片 + VL模型分析文字逻辑关系
    返回融合后的.md文件和分析结果
    """
    global doc_parser
    use_per_page = request.form.get('use_per_page', 'false').lower() == 'true'

    try:
        print("=" * 60)
        print("📤 收到文件上传请求 - 双路径处理模式")
        
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '未找到上传的文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': '文件名为空'}), 400
        
        filename = file.filename
        suffix = Path(filename).suffix.lower()
        
        # 是否启用VL分析（可通过表单参数控制）
        enable_vl = request.form.get('enable_vl', 'true').lower() == 'true'
        
        print(f"📄 文件名: {filename}")
        print(f"📁 文件类型: {suffix}")
        print(f"🤖 VL分析: {'启用' if enable_vl else '禁用'}")
        
        # 检查文件类型
        if suffix not in ['.docx', '.doc', '.pdf']:
            return jsonify({
                'success': False, 
                'error': f'不支持的文件格式: {suffix}，仅支持 .docx, .doc, .pdf'
            }), 400
        
        # 初始化文档解析器（带VL模型配置）
        if doc_parser is None:
            doc_parser = DocumentParser(
                vl_api_key=os.getenv("QWEN_VL_API_KEY", "sk-d7aa5e55b8444f99b49eabb09c0fd263"),
                vl_base_url=os.getenv("QWEN_VL_BASE_URL", "https://bailian.console.aliyun.com/?spm=5176.29619931.J_SEsSjsNv72yRuRFS2VknO.2.74cd10d7Zbu5ig&tab=model#/model-market/all"),
                vl_model=os.getenv("QWEN_VL_MODEL", "Qwen/Qwen3-VL-235B-A22B-Instruct")
            )
        
        # 保存上传的文件到临时目录
        temp_dir = 'D:/存储/Desktop/工作/workplacce/workplacce/TEMP'
        
        # 🔧 修复：确保临时目录存在
        temp_dir_path = Path(temp_dir)
        if not temp_dir_path.exists():
            print(f"📁 创建临时目录: {temp_dir}")
            temp_dir_path.mkdir(parents=True, exist_ok=True)
        
        temp_file_path = temp_dir_path / filename
        
        try:
            file.save(str(temp_file_path))
        except Exception as e:
            print(f"❌ 保存文件失败: {e}")
            # 尝试创建父目录并重试
            temp_file_path.parent.mkdir(parents=True, exist_ok=True)
            file.save(str(temp_file_path))
        
        print(f"💾 已保存到临时目录: {temp_file_path}")
        
        # 流式响应解析进度
        def generate():
            try:
                # 双路径解析文档
                result = doc_parser.parse_document(
                    str(temp_file_path),
                    enable_vl_analysis=enable_vl,
                    use_per_page=True
                )
                
                if result['success']:
                    # 读取生成的融合后.md文件内容
                    md_path = Path(result['md_path'])
                    with open(md_path, 'r', encoding='utf-8') as f:
                        md_content = f.read()
                    
                    # 统计VL分析中发现的特殊布局
                    vl_analysis = result.get('vl_analysis', [])
                    special_layouts = [r for r in vl_analysis if r.get('has_special_layout')]
                    
                    yield json.dumps({
                        'type': 'complete',
                        'success': True,
                        'doc_name': result['doc_name'],
                        'md_content': md_content,
                        'md_path': result['md_path'],
                        'images_dir': result.get('images_dir', ''),
                        'page_images_dir': result.get('page_images_dir', ''),
                        'image_count': result['image_count'],
                        'output_dir': result['output_dir'],
                        'vl_special_layouts': len(special_layouts),
                        'vl_analysis_summary': [
                            {
                                'page': r['page_num'],
                                'type': r.get('layout_type', ''),
                                'has_special': r.get('has_special_layout', False)
                            }
                            for r in vl_analysis
                        ]
                    }) + "\n"
                else:
                    yield json.dumps({
                        'type': 'error',
                        'success': False,
                        'error': result.get('error', '解析失败')
                    }) + "\n"
                    
            except Exception as e:
                import traceback
                traceback.print_exc()
                yield json.dumps({
                    'type': 'error',
                    'success': False,
                    'error': str(e)
                }) + "\n"
            finally:
                # 清理临时文件
                try:
                    if os.path.exists(temp_dir):
                        #shutil.rmtree(temp_dir)
                        print(f"🗑️ 已清理临时目录: {temp_dir}")
                except Exception as e:
                    print(f"⚠️ 清理临时目录失败: {e}")
        
        return Response(
            stream_with_context(generate()),
            mimetype='application/x-ndjson'
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/check_parsed', methods=['POST'])
def check_parsed_document():
    """
    检测已解析的文档（从解析结果中直接检测）
    接收：md_path, images_dir
    """
    try:
        print("=" * 60)
        print("📥 收到解析文档检测请求")
        
        data = request.get_json()
        md_path = data.get('md_path')
        images_dir = data.get('images_dir')
        
        if not md_path or not Path(md_path).exists():
            return jsonify({'success': False, 'error': f'Markdown文件不存在: {md_path}'}), 400
        
        print(f"📝 Markdown路径: {md_path}")
        print(f"🖼️  图片目录: {images_dir}")
        
        # 读取md文件内容
        with open(md_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        filename = Path(md_path).name
        
        # 调用现有的检测流程
        return check_document_internal(content, filename, images_dir)
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


def check_document_internal(content, filename, image_folder):
    """内部检测函数，供API调用"""
    # 流式响应
    def generate():
        try:
            temp_dir = './TEMP'
            temp_file_path = Path(temp_dir) / filename

            # 🔧 修复：确保临时目录存在
            temp_dir_path = Path(temp_dir)
            if not temp_dir_path.exists():
                print(f"📁 创建临时目录: {temp_dir}")
                temp_dir_path.mkdir(parents=True, exist_ok=True)
            
            temp_file_path = temp_dir_path / filename

            # 保存内容到临时文件
            with open(temp_file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print(f"✅ 临时文件已保存: {temp_file_path}")
            
            yield json.dumps({"type": "progress", "message": "正在解析文档...", "percent": 5}) + "\n"
            
            # 解析Markdown
            parser = MarkdownParser(str(temp_file_path), str(image_folder) if image_folder else None)
            text_segments, image_segments = parser.parse()
            
            print(f"📝 文本段落: {len(text_segments)} 个")
            print(f"🖼️  图片: {len(image_segments)} 个")
            
            # 调试输出图片路径
            if image_segments:
                print(f"🔍 调试 - 图片路径:")
                for i, seg in enumerate(image_segments, 1):
                    exists = os.path.exists(seg.image_path)
                    print(f"   {i}. {'✅' if exists else '❌'} {seg.image_path}")
            
            yield json.dumps({
                "type": "progress",
                "message": f"已解析文档，共 {len(text_segments)} 个文本段落，{len(image_segments)} 张图片",
                "percent": 20
            }) + "\n"
            
            all_results = []
            violations = []
            
            total_steps = len(text_segments) + len(image_segments)
            current_step = 0
            
            print(f"\n🔍 开始检测文本内容...")
            
            # 文本检测
            for idx, seg in enumerate(text_segments):
                current_step += 1
                progress = 20 + int((current_step / total_steps) * 70)
                
                msg = f"检测文本段落 {idx+1}/{len(text_segments)}..."
                print(f"  {msg}", end='', flush=True)
                
                yield json.dumps({
                    "type": "progress", 
                    "message": msg, 
                    "percent": progress
                }) + "\n"
                
                # 重试逻辑
                result = None
                max_retries = 5
                for retry in range(max_retries):
                    try:
                        result = text_checker.check_segment(seg)
                        print(f" {'❌' if not result.is_compliant else '✅'}")
                        break
                    except Exception as e:
                        error_msg = str(e)
                        error_type = type(e).__name__
                        is_503_error = "503" in error_msg or "Service is too busy" in error_msg
                        
                        if retry < max_retries - 1:
                            wait_time = 2 ** retry
                            if is_503_error:
                                print(f" ⏳ (服务繁忙，{wait_time}秒后重试...)", end='', flush=True)
                            else:
                                print(f" ⚠️ ({error_type}: {error_msg[:50]}，{wait_time}秒后重试...)", end='', flush=True)
                            time.sleep(wait_time)
                        else:
                            print(f" ❌ (失败: {error_msg[:100]})")
                            from compliance_checker import ViolationResult
                            result = ViolationResult(
                                segment_type="text",
                                paragraph_index=seg.paragraph_index,
                                page_number=seg.page_number,
                                line_range=f"{seg.line_start}-{seg.line_end}",
                                content_preview=seg.content[:50],
                                is_compliant=True,
                                violated_rules=[], severity="low", suggestions="", legal_basis=""
                            )
                
                all_results.append(result)
                if not result.is_compliant:
                    violations.append({
                        'type': 'text',
                        'page_number': result.page_number,
                        'paragraph_index': result.paragraph_index,
                        'line_range': result.line_range,
                        'content': result.content_preview,
                        'severity': result.severity,
                        'violated_rules': result.violated_rules,
                        'suggestions': result.suggestions,
                        'legal_basis': result.legal_basis,
                        'segment_data': {'content': seg.content, 'line_start': seg.line_start, 'line_end': seg.line_end}
                    })

            # 图片检测
            if image_segments and image_checker:
                print(f"\n📷 开始检测图片...")
                for idx, seg in enumerate(image_segments):
                    current_step += 1
                    progress = 20 + int((current_step / total_steps) * 70)
                    
                    msg = f"检测图片 {idx+1}/{len(image_segments)}..."
                    print(f"  {msg}", end='', flush=True)
                    
                    yield json.dumps({
                        "type": "progress", 
                        "message": msg, 
                        "percent": progress
                    }) + "\n"
                    
                    # 重试逻辑
                    result = None
                    max_retries = 5
                    for retry in range(max_retries):
                        try:
                            result = image_checker.check_segment(seg)
                            print(f" {'❌' if not result.is_compliant else '✅'}")
                            break
                        except Exception as e:
                            error_msg = str(e)
                            error_type = type(e).__name__
                            is_503_error = "503" in error_msg or "Service is too busy" in error_msg
                            
                            if retry < max_retries - 1:
                                wait_time = 2 ** retry
                                if is_503_error:
                                    print(f" ⏳ (服务繁忙，{wait_time}秒后重试...)", end='', flush=True)
                                else:
                                    print(f" ⚠️ ({error_type}: {error_msg[:50]}，{wait_time}秒后重试...)", end='', flush=True)
                                time.sleep(wait_time)
                            else:
                                print(f" ❌ (失败: {error_msg[:100]})")
                                from compliance_checker import ViolationResult
                                result = ViolationResult(
                                    segment_type="image",
                                    paragraph_index=seg.paragraph_index,
                                    page_number=seg.page_number,
                                    line_range=seg.line_range,
                                    content_preview=f"图片路径: {seg.image_path}",
                                    is_compliant=True,
                                    violated_rules=[], severity="low", suggestions="", legal_basis=""
                                )
                    
                    all_results.append(result)
                    if not result.is_compliant:
                        violations.append({
                            'type': 'image',
                            'page_number': result.page_number,
                            'paragraph_index': result.paragraph_index,
                            'line_range': result.line_range,
                            'content': f"[图片] {result.content_preview}",
                            'severity': result.severity,
                            'violated_rules': result.violated_rules,
                            'suggestions': result.suggestions,
                            'legal_basis': result.legal_basis,
                            'image_path': seg.image_path
                        })
            
            # 生成报告
            print(f"\n📊 生成最终报告...")
            yield json.dumps({"type": "progress", "message": "正在生成最终报告...", "percent": 95}) + "\n"
            
            risk_counts = {"high": 0, "medium": 0, "low": 0}
            for v in violations:
                severity = v.get('severity', 'low').lower()
                if severity in risk_counts:
                    risk_counts[severity] += 1
            
            total_violations = len(violations)
            if total_violations == 0:
                overall_score = 100
                risk_level = "low"
            else:
                penalty = risk_counts['high'] * 15 + risk_counts['medium'] * 8 + risk_counts['low'] * 3
                overall_score = max(0, 100 - penalty)
                if risk_counts['high'] > 0:
                    risk_level = "high"
                elif risk_counts['medium'] > 2:
                    risk_level = "high"
                elif risk_counts['medium'] > 0:
                    risk_level = "medium"
                else:
                    risk_level = "low"
            
            report = {
                'document_name': filename,
                'check_time': __import__('datetime').datetime.now().isoformat(),
                'overall_score': overall_score,
                'risk_level': risk_level,
                'statistics': {
                    'total_paragraphs': len(text_segments),
                    'total_images': len(image_segments),
                    'total_violations': total_violations,
                    'risk_counts': risk_counts
                },
                'violations': violations
            }
            
            print(f"\n✅ 检测完成!")
            print(f"📊 总分: {overall_score}")
            print(f"⚠️  违规数: {total_violations}")
            
            yield json.dumps({
                "type": "complete",
                "report": report
            }) + "\n"
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            yield json.dumps({
                "type": "error",
                "error": str(e)
            }) + "\n"
        finally:
            # 清理临时文件
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    print(f"🗑️ 已清理临时目录: {temp_dir}")
            except Exception as e:
                print(f"⚠️ 清理临时目录失败: {e}")
    
    return Response(
        stream_with_context(generate()),
        mimetype='application/x-ndjson'
    )


@app.route('/api/check', methods=['POST'])
def check_document():
    """
    检测文档合规性（流式响应）
    """
    try:
        print("=" * 60)
        print("📥 收到检测请求")
        
        data = request.get_json()
        content = data.get('content')
        filename = data.get('filename')
        image_folder = data.get('image_folder', 'images')
        
        print(f"📂 图片文件夹: {image_folder}")
        print(f"📄 文件名: {filename}")
        
        if not content:
            return jsonify({'error': '未提供文件内容'}), 400
        
        # 创建临时目录
        temp_dir = './TEMP'
        temp_file_path = os.path.join(temp_dir, filename)
        
        # 🔧 修复：确保临时目录存在
        temp_dir_path = Path(temp_dir)
        if not temp_dir_path.exists():
            print(f"📁 创建临时目录: {temp_dir}")
            temp_dir_path.mkdir(parents=True, exist_ok=True)
        
        temp_file_path = os.path.join(temp_dir, filename)

        # 修复换行符问题
        if content.count('\n\n') > content.count('\n') * 0.4:
            print("⚠️ 检测到换行符异常，正在标准化...")
            content = re.sub(r'\n{3,}', '\n\n', content)
            content = content.replace('\r\n', '\n')

        # 保存文件
        with open(temp_file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        file_path = temp_file_path
        
        # 强制重新加载markdown_parser
        import importlib
        import markdown_parser
        importlib.reload(markdown_parser)
        
        def generate():
            try:
                # 1. 解析文档
                print("🔍 开始解析文档...")
                yield json.dumps({"type": "progress", "message": "正在解析文档结构...", "percent": 10}) + "\n"
                
                parser = markdown_parser.MarkdownParser(file_path, image_folder)
                text_segments, image_segments = parser.parse()
                
                print(f"📝 文本段落: {len(text_segments)} 个")
                print(f"🖼️  图片: {len(image_segments)} 个")
                
                # 打印图片路径供调试
                if image_segments:
                    print(f"🔍 调试 - 图片路径:")
                    for idx, img_seg in enumerate(image_segments, 1):
                        exists = "✅" if os.path.exists(img_seg.image_path) else "❌"
                        print(f"   {idx}. {exists} {img_seg.image_path}")
                
                yield json.dumps({
                    "type": "progress", 
                    "message": f"解析完成: {len(text_segments)}段文本, {len(image_segments)}张图片", 
                    "percent": 20
                }) + "\n"
                
                # 2. 执行检测
                all_results = []
                violations = []
                
                total_steps = len(text_segments) + len(image_segments)
                current_step = 0
                
                print(f"\n🔍 开始检测文本内容...")
                
                # 文本检测
                for idx, seg in enumerate(text_segments):
                    current_step += 1
                    progress = 20 + int((current_step / total_steps) * 70)
                    
                    msg = f"检测文本段落 {idx+1}/{len(text_segments)}..."
                    print(f"  {msg}", end='', flush=True)
                    
                    yield json.dumps({
                        "type": "progress", 
                        "message": msg, 
                        "percent": progress
                    }) + "\n"
                    
                    # 重试逻辑（增强版：最多5次，指数退避）
                    result = None
                    max_retries = 5
                    for retry in range(max_retries):
                        try:
                            result = text_checker.check_segment(seg)
                            print(f" {'❌' if not result.is_compliant else '✅'}")
                            break
                        except Exception as e:
                            error_msg = str(e)
                            error_type = type(e).__name__
                            # 检查是否是503服务过载错误
                            is_503_error = "503" in error_msg or "Service is too busy" in error_msg or "service_unavailable" in error_msg
                            
                            if retry < max_retries - 1:
                                # 指数退避：1s, 2s, 4s, 8s
                                wait_time = 2 ** retry
                                if is_503_error:
                                    print(f" ⏳ (服务繁忙，{wait_time}秒后重试...)", end='', flush=True)
                                else:
                                    print(f" ⚠️ ({error_type}: {error_msg[:50]}，{wait_time}秒后重试...)", end='', flush=True)
                                time.sleep(wait_time)
                            else:
                                # 最后一次重试失败
                                if is_503_error:
                                    print(f" ❌ (服务持续繁忙，跳过)")
                                else:
                                    print(f" ❌ (失败: {error_msg[:100]})")
                                # 构造失败结果
                                from compliance_checker import ViolationResult
                                result = ViolationResult(
                                    segment_type="text",
                                    paragraph_index=seg.paragraph_index,
                                    page_number=seg.page_number,
                                    line_range=f"{seg.line_start}-{seg.line_end}",
                                    content_preview=seg.content[:50],
                                    is_compliant=True, # 失败默认视为合规以免阻塞
                                    violated_rules=[], severity="low", suggestions="", legal_basis=""
                                )
                    
                    all_results.append(result)
                    if not result.is_compliant:
                        violations.append({
                            'type': 'text',
                            'page_number': result.page_number,
                            'paragraph_index': result.paragraph_index,
                            'line_range': result.line_range,
                            'content': result.content_preview,
                            'severity': result.severity,
                            'violated_rules': result.violated_rules,
                            'suggestions': result.suggestions,
                            'legal_basis': result.legal_basis,
                            'segment_data': {'content': seg.content, 'line_start': seg.line_start, 'line_end': seg.line_end}
                        })

                # 图片检测
                if image_segments and image_checker:
                    print(f"\n📷 开始检测图片...")
                    for idx, seg in enumerate(image_segments):
                        current_step += 1
                        progress = 20 + int((current_step / total_steps) * 70)
                        
                        msg = f"检测图片 {idx+1}/{len(image_segments)}..."
                        print(f"  {msg}", end='', flush=True)
                        
                        yield json.dumps({
                            "type": "progress", 
                            "message": msg, 
                            "percent": progress
                        }) + "\n"
                        
                        # 重试逻辑（增强版：最多5次，指数退避）
                        result = None
                        max_retries = 5
                        for retry in range(max_retries):
                            try:
                                result = image_checker.check_segment(seg)
                                print(f" {'❌' if not result.is_compliant else '✅'}")
                                break
                            except Exception as e:
                                error_msg = str(e)
                                error_type = type(e).__name__
                                # 检查是否是503服务过载错误
                                is_503_error = "503" in error_msg or "Service is too busy" in error_msg or "service_unavailable" in error_msg
                                
                                if retry < max_retries - 1:
                                    # 指数退避：1s, 2s, 4s, 8s
                                    wait_time = 2 ** retry
                                    if is_503_error:
                                        print(f" ⏳ (服务繁忙，{wait_time}秒后重试...)", end='', flush=True)
                                    else:
                                        print(f" ⚠️ ({error_type}: {error_msg[:50]}，{wait_time}秒后重试...)", end='', flush=True)
                                    time.sleep(wait_time)
                                else:
                                    # 最后一次重试失败
                                    if is_503_error:
                                        print(f" ❌ (服务持续繁忙，跳过)")
                                    else:
                                        print(f" ❌ (失败: {error_msg[:100]})")
                                        import traceback
                                        traceback.print_exc()  # 打印完整的错误堆栈
                                    # 构造失败结果
                                    from compliance_checker import ViolationResult
                                    result = ViolationResult(
                                        segment_type="image",
                                        paragraph_index=seg.paragraph_index,
                                        page_number=seg.page_number,
                                        line_range=seg.line_range,
                                        content_preview=f"图片路径: {seg.image_path}",
                                        is_compliant=True, # 失败默认视为合规以免阻塞
                                        violated_rules=[], severity="low", suggestions="", legal_basis=""
                                    )
                        
                        all_results.append(result)
                        if not result.is_compliant:
                            violations.append({
                                'type': 'image',
                                'page_number': result.page_number,
                                'paragraph_index': result.paragraph_index,
                                'line_range': result.line_range,
                                'content': f"[图片] {result.content_preview}",
                                'severity': result.severity,
                                'violated_rules': result.violated_rules,
                                'suggestions': result.suggestions,
                                'legal_basis': result.legal_basis,
                                'image_path': seg.image_path
                            })
                
                # 3. 生成报告
                print(f"\n📊 生成最终报告...")
                yield json.dumps({"type": "progress", "message": "正在生成最终报告...", "percent": 95}) + "\n"
                
                # 计算统计数据
                risk_counts = {"high": 0, "medium": 0, "low": 0}
                for v in violations:
                    severity = v.get('severity', 'low').lower()
                    if severity in risk_counts:
                        risk_counts[severity] += 1
                
                # 计算总分和风险等级
                deductions = {"high": 10, "medium": 5, "low": 2}
                total_deduction = sum(risk_counts[level] * deductions[level] for level in risk_counts)
                overall_score = max(0, 100 - total_deduction)
                
                if risk_counts['high'] > 0:
                    risk_level = 'high'
                elif risk_counts['medium'] > 0:
                    risk_level = 'medium'
                else:
                    risk_level = 'low'
                
                # 构建报告数据
                report_data = {
                    'filename': os.path.basename(file_path),
                    'total_violations': len(violations),
                    'risk_level_count': risk_counts,
                    'high_severity_count': risk_counts['high'],
                    'medium_severity_count': risk_counts['medium'],
                    'low_severity_count': risk_counts['low'],
                    'overall_score': overall_score,
                    'risk_level': risk_level
                }
                
                # 读取原始内容
                with open(file_path, 'r', encoding='utf-8') as f:
                    original_content = f.read()
                
                # 最终结果
                final_result = {
                    'success': True,
                    'original_content': original_content,
                    'violations': violations,
                    'report': report_data,
                    'text_segments': [{'content': s.content, 'line_start': s.line_start, 'line_end': s.line_end, 'paragraph_index': s.paragraph_index, 'page_number': s.page_number} for s in text_segments],
                    'image_segments': [{'image_path': s.image_path, 'alt_text': s.alt_text, 'line_number': s.line_number, 'paragraph_index': s.paragraph_index, 'page_number': s.page_number} for s in image_segments]
                }
                
                print(f"\n✅ 检测完成!")
                print(f"   ❌ 发现违规: {len(violations)} 处")
                sys.stdout.flush()
                
                yield json.dumps({"type": "result", "data": final_result}) + "\n"
                
            except Exception as e:
                import traceback
                traceback.print_exc()
                yield json.dumps({"type": "error", "message": str(e)}) + "\n"
            
            finally:
                # 清理临时文件
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir)
                        print(f"🗑️  已清理临时目录: {temp_dir}")
                    except Exception as e:
                        print(f"⚠️  清理临时文件失败: {e}")

        return Response(stream_with_context(generate()), mimetype='application/x-ndjson')
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/list_files', methods=['POST'])
def list_files():
    """
    列出指定目录下的.md文件
    
    请求参数:
    - directory: 目录路径
    """
    try:
        data = request.get_json()
        directory = data.get('directory', '.')
        
        if not os.path.exists(directory):
            return jsonify({'error': f'目录不存在: {directory}'}), 404
        
        md_files = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith('.md'):
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, directory)
                    md_files.append({
                        'name': file,
                        'path': full_path,
                        'relative_path': rel_path,
                        'size': os.path.getsize(full_path)
                    })
        
        return jsonify({
            'success': True,
            'files': md_files,
            'count': len(md_files)
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/rules', methods=['GET'])
def get_rules():
    """获取合规规则列表"""
    try:
        with open('compliance_rules.json', 'r', encoding='utf-8') as f:
            rules_data = json.load(f)
        
        return jsonify({
            'success': True,
            'rules': rules_data['rules']
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("=" * 70)
    print("教案合规检测 Web 应用")
    print("=" * 70)
    
    # 初始化检测器
    if initialize_checkers():
        print("\n✅ 检测器初始化成功")
        print(f"\n🌐 访问地址: http://localhost:5000")
        print(f"📝 请在浏览器中打开上述地址使用Web界面")
        print("\n按 Ctrl+C 停止服务器")
        print("=" * 70)
        
        # 启动Flask服务器
        app.run(host='0.0.0.0', port=5000, debug=True)
    else:
        print("\n❌ 检测器初始化失败")
        print("请检查 .env 文件中的API配置")
        sys.exit(1)
