#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
教案合规性检测智能体
使用LangChain + DeepSeek + Qwen-VL实现
"""
import os
import sys
import argparse
from pathlib import Path
from typing import List
from dotenv import load_dotenv

from markdown_parser import MarkdownParser, TextSegment, ImageSegment
from compliance_checker import (
    ComplianceRulesLoader,
    TextComplianceChecker,
    ImageComplianceChecker,
    ViolationResult
)
from report_generator import ReportGenerator


class ComplianceAgent:
    """教案合规性检测智能体"""
    
    def __init__(self, 
                 deepseek_api_key: str,
                 deepseek_base_url: str,
                 qwen_vl_api_key: str,
                 qwen_vl_base_url: str,
                 qwen_vl_model: str = None,
                 rules_file: str = "compliance_rules.json"):
        """
        初始化智能体
        
        Args:
            deepseek_api_key: DeepSeek API密钥
            deepseek_base_url: DeepSeek API基础URL
            qwen_vl_api_key: Qwen-VL API密钥
            qwen_vl_base_url: Qwen-VL API基础URL
            qwen_vl_model: Qwen-VL模型名称
            rules_file: 合规规则文件路径
        """
        # 加载合规规则
        self.rules = ComplianceRulesLoader(rules_file)
        
        # 初始化检测器
        self.text_checker = TextComplianceChecker(
            api_key=deepseek_api_key,
            base_url=deepseek_base_url,
            rules=self.rules
        )
        
        self.image_checker = ImageComplianceChecker(
            api_key=qwen_vl_api_key,
            base_url=qwen_vl_base_url,
            rules=self.rules,
            model_name=qwen_vl_model
        )
    
    def check_document(self, markdown_file: str, image_folder: str = "image") -> List[ViolationResult]:
        """
        检测单个markdown文档
        
        Args:
            markdown_file: markdown文件路径
            image_folder: 图片文件夹名称
            
        Returns:
            违规检测结果列表
        """
        print(f"\n📄 开始检测文档: {markdown_file}")
        print("=" * 60)
        
        # 解析markdown文档
        print("\n🔍 步骤 1: 解析markdown文档...")
        parser = MarkdownParser(markdown_file, image_folder)
        text_segments, image_segments = parser.parse()
        
        print(f"   ✓ 发现 {len(text_segments)} 个文本段落")
        print(f"   ✓ 发现 {len(image_segments)} 个图片")
        
        # 检测文本段落
        print("\n🔍 步骤 2: 检测文本内容合规性...")
        all_violations = []
        
        for idx, segment in enumerate(text_segments, 1):
            print(f"   检测文本段落 {idx}/{len(text_segments)}...", end="\r")
            try:
                result = self.text_checker.check_segment(segment)
                all_violations.append(result)
            except Exception as e:
                print(f"\n   ⚠ 文本段落 {idx} 检测出错: {str(e)}")
        
        print(f"\n   ✓ 文本内容检测完成")
        
        # 检测图片内容
        if image_segments:
            print("\n🔍 步骤 3: 检测图片内容合规性...")
            for idx, segment in enumerate(image_segments, 1):
                print(f"   检测图片 {idx}/{len(image_segments)}...", end="\r")
                try:
                    result = self.image_checker.check_segment(segment)
                    all_violations.append(result)
                except Exception as e:
                    print(f"\n   ⚠ 图片 {idx} 检测出错: {str(e)}")
            
            print(f"\n   ✓ 图片内容检测完成")
        else:
            print("\n🔍 步骤 3: 未发现图片，跳过图片检测")
        
        print("\n" + "=" * 60)
        print("✅ 文档检测完成\n")
        
        return all_violations
    
    def check_folder(self, folder_path: str, image_folder: str = "image") -> dict:
        """
        检测文件夹中的所有markdown文档
        
        Args:
            folder_path: 文件夹路径
            image_folder: 图片文件夹名称
            
        Returns:
            每个文档的检测结果字典 {文档名: 违规结果列表}
        """
        folder = Path(folder_path)
        markdown_files = list(folder.glob("*.md")) + list(folder.glob("*.markdown"))
        
        if not markdown_files:
            print(f"⚠ 在 {folder_path} 中未找到markdown文件")
            return {}
        
        print(f"\n📁 发现 {len(markdown_files)} 个markdown文件")
        print("=" * 60)
        
        results = {}
        for md_file in markdown_files:
            violations = self.check_document(str(md_file), image_folder)
            results[md_file.name] = violations
        
        return results
    
    def generate_report(self, document_name: str, violations: List[ViolationResult],
                       output_dir: str = "reports"):
        """
        生成检测报告
        
        Args:
            document_name: 文档名称
            violations: 违规检测结果列表
            output_dir: 报告输出目录
        """
        print(f"\n📊 生成检测报告...")
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 生成报告
        generator = ReportGenerator(document_name)
        report = generator.generate_report(violations)
        
        # 保存报告
        base_name = Path(document_name).stem
        json_path = os.path.join(output_dir, f"{base_name}_report1 .json")
        md_path = os.path.join(output_dir, f"{base_name}_report1.md")
        
        generator.save_report_json(report, json_path)
        generator.save_report_markdown(report, md_path)
        
        print(f"   ✓ JSON报告已保存: {json_path}")
        print(f"   ✓ Markdown报告已保存: {md_path}")
        
        # 打印摘要
        print("\n" + "=" * 60)
        print("📈 检测结果摘要")
        print("=" * 60)
        print(f"   文档名称: {document_name}")
        print(f"   总体得分: {report['总体得分']}/100")
        print(f"   风险等级: {report['风险等级']}")
        print(f"   违规段落: {report['检测统计']['违规段落数']}处")
        print(f"   合规率: {report['检测统计']['合规率']}")
        print("=" * 60 + "\n")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="教案合规性检测智能体",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  # 检测单个文档
  python agent.py --file lesson.md
  
  # 检测整个文件夹
  python agent.py --folder ./lessons
  
  # 指定图片文件夹名称
  python agent.py --file lesson.md --image-folder images
  
  # 指定报告输出目录
  python agent.py --file lesson.md --output reports
        """
    )
    
    parser.add_argument(
        "--file",
        type=str,
        help="要检测的markdown文件路径"
    )
    
    parser.add_argument(
        "--folder",
        type=str,
        help="要检测的文件夹路径（批量检测）"
    )
    
    parser.add_argument(
        "--image-folder",
        type=str,
        default="image",
        help="图片文件夹名称（默认: image）"
    )
    
    parser.add_argument(
        "--output",
        type=str,
        default="reports",
        help="报告输出目录（默认: reports）"
    )
    
    parser.add_argument(
        "--rules",
        type=str,
        default="compliance_rules.json",
        help="合规规则文件路径（默认: compliance_rules.json）"
    )
    
    args = parser.parse_args()
    
    # 检查参数
    if not args.file and not args.folder:
        parser.print_help()
        print("\n❌ 错误: 请指定 --file 或 --folder 参数")
        sys.exit(1)
    
    # 加载环境变量
    load_dotenv()
    
    deepseek_api_key = os.getenv("DEEPSEEK_API_KEY")
    deepseek_base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
    # Qwen-VL服务不需要API密钥，使用默认值
    qwen_vl_api_key = os.getenv("QWEN_VL_API_KEY", "EMPTY")
    qwen_vl_base_url = os.getenv("QWEN_VL_BASE_URL", "http://xn-d.suanjiayun.com:57915/v1")
    qwen_vl_model = os.getenv("QWEN_VL_MODEL", "Qwen/Qwen3-VL-235B-A22B-Instruct-FP8")
    
    # 检查API密钥
    if not deepseek_api_key:
        print("❌ 错误: 未设置 DEEPSEEK_API_KEY 环境变量")
        print("请在 .env 文件中配置API密钥")
        sys.exit(1)
    
    # Qwen-VL不需要检查API密钥
    
    # 初始化智能体
    print("\n🤖 初始化教案合规性检测智能体...")
    print(f"   - DeepSeek Base URL: {deepseek_base_url}")
    print(f"   - Qwen-VL Base URL: {qwen_vl_base_url}")
    print(f"   - Qwen-VL Model: {qwen_vl_model}")
    agent = ComplianceAgent(
        deepseek_api_key=deepseek_api_key,
        deepseek_base_url=deepseek_base_url,
        qwen_vl_api_key=qwen_vl_api_key,
        qwen_vl_base_url=qwen_vl_base_url,
        qwen_vl_model=qwen_vl_model,
        rules_file=args.rules
    )
    print("✅ 智能体初始化完成")
    
    # 执行检测
    if args.file:
        # 单文件检测
        if not os.path.exists(args.file):
            print(f"❌ 错误: 文件不存在 - {args.file}")
            sys.exit(1)
        
        violations = agent.check_document(args.file, args.image_folder)
        agent.generate_report(args.file, violations, args.output)
        
    elif args.folder:
        # 文件夹批量检测
        if not os.path.isdir(args.folder):
            print(f"❌ 错误: 文件夹不存在 - {args.folder}")
            sys.exit(1)
        
        results = agent.check_folder(args.folder, args.image_folder)
        
        # 为每个文档生成报告
        for doc_name, violations in results.items():
            agent.generate_report(doc_name, violations, args.output)
    
    print("\n🎉 所有检测任务完成！\n")


if __name__ == "__main__":
    main()