"""
检测报告生成器
生成格式化的合规性检测报告
"""
import json
from datetime import datetime
from typing import List, Dict
from compliance_checker import ViolationResult


class ReportGenerator:
    """检测报告生成器"""
    
    def __init__(self, document_name: str):
        """
        初始化报告生成器
        
        Args:
            document_name: 文档名称
        """
        self.document_name = document_name
        self.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def generate_report(self, violations: List[ViolationResult]) -> Dict:
        """
        生成检测报告
        
        Args:
            violations: 违规检测结果列表
            
        Returns:
            报告字典
        """
        # 计算总体得分和风险等级
        total_score = self._calculate_score(violations)
        risk_level = self._calculate_risk_level(violations)
        
        # 统计违规情况
        violation_stats = self._get_violation_statistics(violations)
        
        # 过滤出违规的段落
        violation_details = [
            v for v in violations if not v.is_compliant
        ]
        
        # 构建报告
        report = {
            "document_name": self.document_name,
            "检测时间": self.timestamp,
            "总体得分": total_score,
            "风险等级": risk_level,
            "检测统计": {
                "总段落数": len(violations),
                "文本段落数": sum(1 for v in violations if v.segment_type == "text"),
                "图片段落数": sum(1 for v in violations if v.segment_type == "image"),
                "违规段落数": len(violation_details),
                "合规率": f"{(len(violations) - len(violation_details)) / len(violations) * 100:.2f}%"
            },
            "违规统计": violation_stats,
            "违规段落详情": [self._format_violation(v) for v in violation_details],
            "总体建议": self._generate_overall_suggestions(violations)
        }
        
        return report
    
    def _calculate_score(self, violations: List[ViolationResult]) -> float:
        """
        计算总体得分（100分制）
        
        根据违规的严重程度和数量计算得分
        """
        if not violations:
            return 100.0
        
        total_deduction = 0
        for violation in violations:
            if not violation.is_compliant:
                # 根据严重程度扣分
                if violation.severity == "high":
                    total_deduction += 10
                elif violation.severity == "medium":
                    total_deduction += 5
                elif violation.severity == "low":
                    total_deduction += 2
        
        score = max(0, 100 - total_deduction)
        return round(score, 2)
    
    def _calculate_risk_level(self, violations: List[ViolationResult]) -> str:
        """
        计算风险等级
        
        Returns:
            "低风险", "中风险", "高风险", "极高风险"
        """
        high_count = sum(1 for v in violations if not v.is_compliant and v.severity == "high")
        medium_count = sum(1 for v in violations if not v.is_compliant and v.severity == "medium")
        low_count = sum(1 for v in violations if not v.is_compliant and v.severity == "low")
        
        if high_count >= 3:
            return "极高风险"
        elif high_count >= 1:
            return "高风险"
        elif medium_count >= 3:
            return "高风险"
        elif medium_count >= 1 or low_count >= 3:
            return "中风险"
        elif low_count >= 1:
            return "低风险"
        else:
            return "无风险"
    
    def _get_violation_statistics(self, violations: List[ViolationResult]) -> Dict:
        """获取违规统计信息"""
        stats = {
            "高风险违规": 0,
            "中风险违规": 0,
            "低风险违规": 0,
            "违规类别分布": {}
        }
        
        for violation in violations:
            if not violation.is_compliant:
                # 统计严重程度
                if violation.severity == "high":
                    stats["高风险违规"] += 1
                elif violation.severity == "medium":
                    stats["中风险违规"] += 1
                elif violation.severity == "low":
                    stats["低风险违规"] += 1
                
                # 统计违规类别
                for rule in violation.violated_rules:
                    category = rule.get("category", "未知类别")
                    stats["违规类别分布"][category] = stats["违规类别分布"].get(category, 0) + 1
        
        return stats
    
    def _format_violation(self, violation: ViolationResult) -> Dict:
        """格式化单个违规记录"""
        return {
            "类型": "文本段落" if violation.segment_type == "text" else "图片内容",
            "段落索引": violation.paragraph_index + 1,
            "页码": violation.page_number,
            "行号范围": violation.line_range,
            "内容预览": violation.content_preview,
            "严重程度": violation.severity,
            "违反的规则": [
                f"{rule['rule_name']} (规则{rule['rule_id']})"
                for rule in violation.violated_rules
            ],
            "修改建议": violation.suggestions,
            "法律依据": violation.legal_basis
        }
    
    def _generate_overall_suggestions(self, violations: List[ViolationResult]) -> str:
        """生成总体建议"""
        violation_list = [v for v in violations if not v.is_compliant]
        
        if not violation_list:
            return "该教案内容符合所有合规要求，可以正常使用。"
        
        suggestions = []
        
        # 按严重程度给出建议
        high_violations = [v for v in violation_list if v.severity == "high"]
        if high_violations:
            suggestions.append(
                f"发现{len(high_violations)}处高风险违规内容，建议立即修改。"
                "这些内容可能涉及法律法规问题，必须在使用前完全整改。"
            )
        
        medium_violations = [v for v in violation_list if v.severity == "medium"]
        if medium_violations:
            suggestions.append(
                f"发现{len(medium_violations)}处中风险问题，建议优先处理。"
                "这些内容可能影响教学效果或引发争议。"
            )
        
        low_violations = [v for v in violation_list if v.severity == "low"]
        if low_violations:
            suggestions.append(
                f"发现{len(low_violations)}处低风险问题，建议在后续版本中改进。"
            )
        
        # 给出具体的改进方向
        suggestions.append("\n改进建议：")
        suggestions.append("1. 请仔细阅读每个违规段落的详细说明和修改建议")
        suggestions.append("2. 参考提供的法律依据，确保内容符合相关规定")
        suggestions.append("3. 对于图片内容，可能需要更换或添加适当的说明文字")
        suggestions.append("4. 修改完成后建议重新进行合规性检测")
        
        return "\n".join(suggestions)
    
    def save_report_json(self, report: Dict, output_path: str):
        """
        保存报告为JSON文件
        
        Args:
            report: 报告字典
            output_path: 输出文件路径
        """
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
    
    def save_report_markdown(self, report: Dict, output_path: str):
        """
        保存报告为Markdown文件
        
        Args:
            report: 报告字典
            output_path: 输出文件路径
        """
        md_content = self._generate_markdown(report)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(md_content)
    
    def _generate_markdown(self, report: Dict) -> str:
        """生成Markdown格式的报告"""
        md = f"""# 教案合规性检测报告

## 基本信息

- **文档名称**: {report['document_name']}
- **检测时间**: {report['检测时间']}
- **总体得分**: {report['总体得分']}/100
- **风险等级**: {report['风险等级']}

## 检测统计

| 项目 | 数量/比例 |
|------|----------|
| 总段落数 | {report['检测统计']['总段落数']} |
| 文本段落数 | {report['检测统计']['文本段落数']} |
| 图片段落数 | {report['检测统计']['图片段落数']} |
| 违规段落数 | {report['检测统计']['违规段落数']} |
| 合规率 | {report['检测统计']['合规率']} |

## 违规统计

- **高风险违规**: {report['违规统计']['高风险违规']}处
- **中风险违规**: {report['违规统计']['中风险违规']}处
- **低风险违规**: {report['违规统计']['低风险违规']}处

### 违规类别分布

"""
        for category, count in report['违规统计']['违规类别分布'].items():
            md += f"- {category}: {count}处\n"
        
        md += "\n## 违规段落详情\n\n"
        
        if not report['违规段落详情']:
            md += "✅ 未发现违规内容，该教案符合所有合规要求。\n\n"
        else:
            for idx, violation in enumerate(report['违规段落详情'], 1):
                md += f"### 违规 {idx}：{violation['类型']}\n\n"
                md += f"- **位置**: 第{violation['页码']}页，段落{violation['段落索引']}"
                if violation['行号范围']:
                    md += f"，行号{violation['行号范围']}"
                md += "\n"
                md += f"- **严重程度**: {violation['严重程度']}\n"
                md += f"- **内容预览**: {violation['内容预览']}\n"
                md += f"- **违反的规则**: {', '.join(violation['违反的规则'])}\n"
                md += f"- **修改建议**: {violation['修改建议']}\n"
                if violation['法律依据']:
                    md += f"- **法律依据**: {violation['法律依据']}\n"
                md += "\n---\n\n"
        
        md += "## 总体建议\n\n"
        md += report['总体建议']
        md += "\n\n---\n\n"
        md += "*本报告由AI智能体自动生成，仅供参考。最终解释权归相关监管部门所有。*\n"
        
        return md
