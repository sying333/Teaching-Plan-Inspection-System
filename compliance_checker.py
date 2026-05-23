"""
合规性检测器
使用LangChain集成DeepSeek和Qwen-VL进行文本和图片合规性检测
"""
import os
import json
import base64
from typing import List, Dict, Optional
from dataclasses import dataclass, asdict

from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.messages import HumanMessage

from markdown_parser import TextSegment, ImageSegment, get_page_number_from_line


@dataclass
class ViolationResult:
    """违规检测结果"""
    segment_type: str  # "text" or "image"
    paragraph_index: int
    page_number: int
    line_range: Optional[str]
    content_preview: str
    is_compliant: bool
    violated_rules: List[Dict]  # 违反的规则列表
    severity: str  # "high", "medium", "low"
    suggestions: str  # 修改建议
    legal_basis: str  # 法律依据


class ComplianceRulesLoader:
    """合规规则加载器"""
    
    def __init__(self, rules_file: str = "compliance_rules.json"):
        self.rules_file = rules_file
        self.rules = self._load_rules()
    
    def _load_rules(self) -> List[Dict]:
        """加载合规规则"""
        with open(self.rules_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data['rules']
    
    def get_rules_text(self) -> str:
        """获取规则的文本描述"""
        rules_text = "教案合规性检测规则：\n\n"
        for rule in self.rules:
            rules_text += f"{rule['id']}. {rule['name']}：{rule['description']}\n"
            rules_text += f"   - 严重程度：{rule['severity']}\n"
            rules_text += f"   - 分类：{rule['category']}\n\n"
        return rules_text


class TextComplianceChecker:
    """文本合规性检测器（使用DeepSeek）"""
    
    def __init__(self, api_key: str, base_url: str, rules: ComplianceRulesLoader):
        """
        初始化文本检测器
        
        Args:
            api_key: DeepSeek API密钥
            base_url: DeepSeek API基础URL
            rules: 合规规则加载器
        """
        self.llm = ChatOpenAI(
            model="deepseek-chat",
            openai_api_key=api_key,
            openai_api_base=base_url,
            temperature=0.1,
            max_tokens=2000,
            request_timeout=60  # 60秒超时（增加容错性）
        )
        self.rules = rules
        
    def check_segment(self, segment: TextSegment) -> ViolationResult:
        """
        检测文本段落的合规性
        
        Args:
            segment: 文本段落
            
        Returns:
            检测结果
        """
        # 构建提示词
        prompt = ChatPromptTemplate.from_messages([
            ("system", """你是一位专业的教案合规性审核专家。你需要根据提供的合规规则，仔细审查教案文本内容。

{rules}

请按照以下JSON格式返回检测结果：
{{
    "is_compliant": true/false,
    "violated_text": "具体违规的原文句子或短语（如果合规则为空字符串）",
    "violated_rules": [
        {{
            "rule_id": 规则ID,
            "rule_name": "规则名称",
            "severity": "high/medium/low",
            "category": "分类"
        }}
    ],
    "overall_severity": "high/medium/low/none",
    "suggestions": "具体的修改建议",
    "legal_basis": "相关的法律法规依据"
}}

注意：
1. violated_text 必须是原文中的具体违规句子或短语，不要修改原文
2. 如果存在多处违规，只返回最严重的那一处的原文
3. 请严格按照JSON格式返回，不要添加任何额外的文字说明"""),
            ("human", "请检测以下教案文本段落的合规性：\n\n{text}")
        ])
        
        # 调用LLM
        chain = prompt | self.llm
        response = chain.invoke({
            "rules": self.rules.get_rules_text(),
            "text": segment.content
        })
        
        # 解析结果
        try:
            result_dict = json.loads(response.content)
        except json.JSONDecodeError:
            # 如果解析失败，尝试提取JSON部分
            content = response.content
            start_idx = content.find('{')
            end_idx = content.rfind('}') + 1
            if start_idx != -1 and end_idx > start_idx:
                result_dict = json.loads(content[start_idx:end_idx])
            else:
                # 解析失败，返回默认结果
                result_dict = {
                    "is_compliant": True,
                    "violated_rules": [],
                    "overall_severity": "none",
                    "suggestions": "",
                    "legal_basis": ""
                }
        
        # 构建违规结果
        line_range = f"{segment.line_start}-{segment.line_end}"
        
        # 优先使用LLM返回的具体违规文本，否则使用段落前100字符
        violated_text = result_dict.get("violated_text", "")
        if violated_text and len(violated_text) > 5:
            content_preview = violated_text
        else:
            content_preview = segment.content[:100] + "..." if len(segment.content) > 100 else segment.content
        
        return ViolationResult(
            segment_type="text",
            paragraph_index=segment.paragraph_index,
            page_number=segment.page_number,
            line_range=line_range,
            content_preview=content_preview,
            is_compliant=result_dict.get("is_compliant", True),
            violated_rules=result_dict.get("violated_rules", []),
            severity=result_dict.get("overall_severity", "none"),
            suggestions=result_dict.get("suggestions", ""),
            legal_basis=result_dict.get("legal_basis", "")
        )


class ImageComplianceChecker:
    """图片合规性检测器（使用Qwen-VL）"""
    
    def __init__(self, api_key: str, base_url: str, rules: ComplianceRulesLoader, model_name: str = None):
        """
        初始化图片检测器
        
        Args:
            api_key: Qwen-VL API密钥
            base_url: Qwen-VL API基础URL
            rules: 合规规则加载器
            model_name: 模型名称，默认为None时使用环境变量或默认值
        """
        # 如果未指定模型名称，尝试从环境变量读取，否则使用默认值
        if model_name is None:
            model_name = os.getenv("QWEN_VL_MODEL", "Qwen/Qwen3-VL-235B-A22B-Instruct-FP8")
        
        self.llm = ChatOpenAI(
            model=model_name,
            openai_api_key=api_key,
            openai_api_base=base_url,
            temperature=0.1,
            max_tokens=2000,
            request_timeout=60  # 60秒超时（增加容错性）
        )
        self.rules = rules
    
    def _encode_image(self, image_path: str) -> str:
        """将图片编码为base64"""
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    
    def check_segment(self, segment: ImageSegment) -> ViolationResult:
        """
        检测图片及其上下文的合规性
        
        Args:
            segment: 图片段落
            
        Returns:
            检测结果
        """
        # 检查图片文件是否存在
        if not os.path.exists(segment.image_path):
            return ViolationResult(
                segment_type="image",
                paragraph_index=segment.paragraph_index,
                page_number=segment.page_number,
                line_range=str(segment.line_number),
                content_preview=f"图片未找到: {segment.image_path}",
                is_compliant=False,
                violated_rules=[{
                    "rule_id": 0,
                    "rule_name": "资源完整性",
                    "severity": "high",
                    "category": "技术问题"
                }],
                severity="high",
                suggestions="请确保图片文件存在于正确的路径",
                legal_basis=""
            )
        
        # 编码图片
        try:
            base64_image = self._encode_image(segment.image_path)
        except Exception as e:
            return ViolationResult(
                segment_type="image",
                paragraph_index=segment.paragraph_index,
                page_number=segment.page_number,
                line_range=str(segment.line_number),
                content_preview=f"图片读取错误: {str(e)}",
                is_compliant=False,
                violated_rules=[{
                    "rule_id": 0,
                    "rule_name": "资源完整性",
                    "severity": "high",
                    "category": "技术问题"
                }],
                severity="high",
                suggestions="请检查图片文件是否损坏",
                legal_basis=""
            )
        
        # 构建提示词
        context_text = f"""
图片描述：{segment.alt_text}
图片前的上下文：{segment.context_before}
图片后的上下文：{segment.context_after}
"""
        
        prompt_text = f"""你是一位专业的教案合规性审核专家。你需要根据提供的合规规则，仔细审查教案中的图片内容及其上下文。

{self.rules.get_rules_text()}

请分析以下图片及其上下文信息：
{context_text}

请按照以下JSON格式返回检测结果：
{{
    "is_compliant": true/false,
    "violated_rules": [
        {{
            "rule_id": 规则ID,
            "rule_name": "规则名称",
            "severity": "high/medium/low",
            "category": "分类"
        }}
    ],
    "overall_severity": "high/medium/low/none",
    "suggestions": "具体的修改建议",
    "legal_basis": "相关的法律法规依据"
}}

请严格按照JSON格式返回，不要添加任何额外的文字说明。"""
        
        # 调用多模态LLM
        try:
            message = HumanMessage(
                content=[
                    {
                        "type": "text",
                        "text": prompt_text
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}"
                        }
                    }
                ]
            )
            
            response = self.llm.invoke([message])
            
            # 解析结果
            try:
                result_dict = json.loads(response.content)
            except json.JSONDecodeError:
                content = response.content
                start_idx = content.find('{')
                end_idx = content.rfind('}') + 1
                if start_idx != -1 and end_idx > start_idx:
                    result_dict = json.loads(content[start_idx:end_idx])
                else:
                    result_dict = {
                        "is_compliant": True,
                        "violated_rules": [],
                        "overall_severity": "none",
                        "suggestions": "",
                        "legal_basis": ""
                    }
        except Exception as e:
            print(f"图片检测出错: {str(e)}")
            result_dict = {
                "is_compliant": True,
                "violated_rules": [],
                "overall_severity": "none",
                "suggestions": "",
                "legal_basis": ""
            }
        
        # 构建违规结果
        content_preview = f"图片: {segment.alt_text} ({os.path.basename(segment.image_path)})"
        
        return ViolationResult(
            segment_type="image",
            paragraph_index=segment.paragraph_index,
            page_number=segment.page_number,
            line_range=str(segment.line_number),
            content_preview=content_preview,
            is_compliant=result_dict.get("is_compliant", True),
            violated_rules=result_dict.get("violated_rules", []),
            severity=result_dict.get("overall_severity", "none"),
            suggestions=result_dict.get("suggestions", ""),
            legal_basis=result_dict.get("legal_basis", "")
        )
