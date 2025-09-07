from openai import OpenAI
import os
import json
import re
from typing import Dict, List, Optional

class QwenAnalysisService:
    """
    基于Qwen大模型的文档分析服务
    """
    
    def __init__(self):
        self.client = OpenAI(
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
        )
        self.model = "qwen-plus-2025-04-28"
    
    def analyze_tender_document(self, content: str) -> Dict:
        """
        分析招标文件，提取废标条款
        
        Args:
            content: 招标文件内容
            
        Returns:
            分析结果字典，包含废标项和建议
        """
        prompt = f"""
        请分析以下招标文件内容，提取出所有可能导致投标无效的条款（废标条款）。
        
        招标文件内容：
        {content}
        
        请按照以下JSON格式返回分析结果：
        {{
            "summary": "招标文件分析摘要",
            "invalid_items": [
                {{
                    "category": "条款类别（如：资质要求、技术要求、商务要求等）",
                    "description": "具体的废标条款描述",
                    "requirement": "具体要求内容",
                    "severity": "严重程度（高/中/低）",
                    "keywords": ["关键词1", "关键词2"]
                }}
            ],
            "suggestions": [
                "投标建议1",
                "投标建议2"
            ]
        }}
        
        重点关注以下方面：
        1. 资质要求（营业执照、资质证书、业绩要求等）
        2. 技术要求（技术规格、性能指标、标准符合性等）
        3. 商务要求（价格、付款方式、交期要求等）
        4. 格式要求（文件格式、签名盖章、装订要求等）
        5. 时间要求（投标截止时间、有效期等）
        6. 其他可能导致废标的条款
        
        请确保返回的是有效的JSON格式。
        """
        
        try:
            response = self._call_qwen_api(prompt)
            return self._parse_tender_response(response)
        except Exception as e:
            return {
                "summary": f"分析过程中出现错误：{str(e)}",
                "invalid_items": [],
                "suggestions": ["请检查文件内容或重新尝试分析"]
            }
    
    def analyze_bid_document(self, content: str, tender_analysis: Optional[Dict] = None) -> Dict:
        """
        分析投标文件，检查是否存在废标风险
        
        Args:
            content: 投标文件内容
            tender_analysis: 招标文件分析结果（可选）
            
        Returns:
            合规性检查结果
        """
        base_prompt = f"""
        请分析以下投标文件内容，检查是否存在可能导致废标的问题。
        
        投标文件内容：
        {content}
        """
        
        if tender_analysis and 'invalid_items' in tender_analysis:
            invalid_items_text = "\n".join([
                f"- {item.get('category', '未知类别')}: {item.get('description', '')}"
                for item in tender_analysis['invalid_items']
            ])
            
            base_prompt += f"""
            
            招标文件中发现的废标条款：
            {invalid_items_text}
            
            请特别关注投标文件是否满足上述要求。
            """
        
        prompt = base_prompt + """
        
        请按照以下JSON格式返回分析结果：
        {
            "summary": "投标文件合规性分析摘要",
            "compliance_check": {
                "overall_status": "合规性状态（合规/存在风险/不合规）",
                "risk_level": "风险等级（低/中/高）",
                "score": "合规性得分（0-100）"
            },
            "issues": [
                {
                    "category": "问题类别",
                    "description": "具体问题描述",
                    "severity": "严重程度（高/中/低）",
                    "suggestion": "改进建议",
                    "location": "问题位置（如果能定位）"
                }
            ],
            "recommendations": [
                "改进建议1",
                "改进建议2"
            ]
        }
        
        检查重点：
        1. 是否满足资质要求
        2. 技术规格是否符合要求
        3. 商务条件是否满足
        4. 文件格式是否正确
        5. 是否有遗漏的必要信息
        6. 签名盖章是否完整
        
        请确保返回的是有效的JSON格式。
        """
        
        try:
            response = self._call_qwen_api(prompt)
            return self._parse_bid_response(response)
        except Exception as e:
            return {
                "summary": f"分析过程中出现错误：{str(e)}",
                "compliance_check": {
                    "overall_status": "未知",
                    "risk_level": "未知",
                    "score": 0
                },
                "issues": [],
                "recommendations": ["请检查文件内容或重新尝试分析"]
            }
    
    def _call_qwen_api(self, prompt: str) -> str:
        """
        调用Qwen API
        """
        messages = [{"role": "user", "content": prompt}]
        
        completion = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            stream=False,
            temperature=0.3  # 降低随机性，提高一致性
        )
        
        return completion.choices[0].message.content
    
    def _parse_tender_response(self, response: str) -> Dict:
        """
        解析招标文件分析响应
        """
        try:
            # 尝试直接解析JSON
            return json.loads(response)
        except json.JSONDecodeError:
            # 如果直接解析失败，尝试提取JSON部分
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                try:
                    return json.loads(json_match.group())
                except json.JSONDecodeError:
                    pass
            
            # 如果无法解析，返回默认格式
            return {
                "summary": "解析响应时出现错误，但分析已完成",
                "invalid_items": self._extract_invalid_items_from_text(response),
                "suggestions": self._extract_suggestions_from_text(response)
            }
    
    def _parse_bid_response(self, response: str) -> Dict:
        """
        解析投标文件分析响应
        """
        try:
            # 尝试直接解析JSON
            return json.loads(response)
        except json.JSONDecodeError:
            # 如果直接解析失败，尝试提取JSON部分
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                try:
                    return json.loads(json_match.group())
                except json.JSONDecodeError:
                    pass
            
            # 如果无法解析，返回默认格式
            return {
                "summary": "解析响应时出现错误，但分析已完成",
                "compliance_check": {
                    "overall_status": "需要人工审核",
                    "risk_level": "中",
                    "score": 50
                },
                "issues": self._extract_issues_from_text(response),
                "recommendations": self._extract_suggestions_from_text(response)
            }
    
    def _extract_invalid_items_from_text(self, text: str) -> List[Dict]:
        """
        从文本中提取废标项信息
        """
        items = []
        # 简单的文本解析逻辑
        lines = text.split('\n')
        for line in lines:
            if any(keyword in line for keyword in ['废标', '无效', '不符合', '要求']):
                items.append({
                    "category": "从文本提取",
                    "description": line.strip(),
                    "requirement": "",
                    "severity": "中",
                    "keywords": []
                })
        return items[:10]  # 限制数量
    
    def _extract_issues_from_text(self, text: str) -> List[Dict]:
        """
        从文本中提取问题信息
        """
        issues = []
        lines = text.split('\n')
        for line in lines:
            if any(keyword in line for keyword in ['问题', '缺少', '不完整', '错误']):
                issues.append({
                    "category": "从文本提取",
                    "description": line.strip(),
                    "severity": "中",
                    "suggestion": "请核实相关内容",
                    "location": "未定位"
                })
        return issues[:10]  # 限制数量
    
    def _extract_suggestions_from_text(self, text: str) -> List[str]:
        """
        从文本中提取建议信息
        """
        suggestions = []
        lines = text.split('\n')
        for line in lines:
            if any(keyword in line for keyword in ['建议', '应该', '需要', '注意']):
                suggestions.append(line.strip())
        return suggestions[:5]  # 限制数量
