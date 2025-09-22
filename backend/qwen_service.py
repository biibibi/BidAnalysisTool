#!/usr/bin/env python3
"""
Qwen AI 文档分析服务
===================

本模块封装了基于阿里云百炼平台Qwen大模型的文档智能分析功能。
主要提供招标文件分析和投标文件合规性检查两大核心功能。

核心功能：
    1. 招标文件分析 - 提取废标条款、资质要求、技术要求等
    2. 投标文件分析 - 检查合规性、评估风险等级、提供改进建议

技术特点：
    - 使用OpenAI SDK兼容模式调用阿里云百炼API
    - 支持流式和非流式响应
    - 智能JSON解析和错误处理
    - 结构化的分析结果输出

依赖库：
    - openai: OpenAI Python SDK
    - os: 环境变量管理
    - json: JSON数据处理
    - re: 正则表达式
    - typing: 类型注解

环境要求：
    - DASHSCOPE_API_KEY: 阿里云百炼API密钥

作者：BidAnalysis Team
创建时间：2025年
版本：1.0
"""

from openai import OpenAI
import os
import json
import re
from typing import Dict, List, Optional, Any, cast
from dotenv import load_dotenv

# 确保在独立脚本/测试环境下也能读取 backend/.env
_env_loaded = False
try:
    _backend_dir = os.path.dirname(os.path.abspath(__file__))
    _env_path = os.path.join(_backend_dir, '.env')
    if os.path.exists(_env_path):
        load_dotenv(_env_path, override=False)
        _env_loaded = True
except Exception:
    # 静默失败，不影响后续从系统环境读取
    _env_loaded = False

class QwenAnalysisService:
    """
    基于Qwen大模型的文档分析服务类
    ================================
    
    该类封装了与阿里云百炼平台Qwen模型的交互逻辑，
    提供招标文件和投标文件的智能分析功能。
    
    主要方法：
        - analyze_tender_document(): 分析招标文件
        - analyze_bid_document(): 分析投标文件
        - _call_qwen_api(): 调用AI API
        - _parse_*_response(): 解析AI响应
    
    使用示例：
        service = QwenAnalysisService()
        result = service.analyze_tender_document(content)
    """
    
    def __init__(self):
        """
        初始化Qwen分析服务
        
        配置OpenAI客户端以连接阿里云百炼平台，
        设置默认模型和API参数。
        
        Raises:
            ValueError: 当API密钥未设置时
        """
        # 初始化OpenAI客户端，使用阿里云百炼的兼容模式
        self.client = OpenAI(
            # 从环境变量获取API密钥，确保安全性
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            # 阿里云百炼平台的API端点
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
        )
        # 使用的模型版本，可根据需要调整
        self.model = "qwen-plus-2025-04-28"
        # 默认大模型提供方（仅作为可选属性，不改变现有行为）
        self.default_provider = os.getenv("LLM_PROVIDER", "qwen").lower()

        # 豆包（字节方舟 Ark）默认配置（不影响现有流程）
        self.ark_base_url = os.getenv("ARK_BASE_URL", "https://ark.cn-beijing.volces.com/api/v3")
        # 模型ID优先取环境变量，若未配置则回退到示例模型名
        # 如您有专属推理接入点ID（UUID），请设置 DOUBAO_MODEL_ID=您的接入点ID
        self.doubao_model_id = os.getenv("DOUBAO_MODEL_ID", "doubao-seed-1-6-250615")
    
    def analyze_tender_document(self, content: str) -> Dict:
        """
        分析招标文件，提取废标条款和关键要求
        ====================================
        
        使用Qwen AI对招标文件内容进行深度分析，
        自动识别可能导致投标无效的各类条款。
        
        Args:
            content (str): 招标文件的文本内容
            
        Returns:
            Dict: 分析结果字典，包含以下字段：
                - summary (str): 招标文件分析摘要
                - invalid_items (List[Dict]): 废标条款列表
                    - category (str): 条款类别
                    - description (str): 具体描述
                    - requirement (str): 具体要求
                    - severity (str): 严重程度(高/中/低)
                    - keywords (List[str]): 关键词列表
                - suggestions (List[str]): 投标建议列表
                
        分析维度：
            1. 资质要求：营业执照、资质证书、业绩要求等
            2. 技术要求：技术规格、性能指标、标准符合性等
            3. 商务要求：价格、付款方式、交期要求等
            4. 格式要求：文件格式、签名盖章、装订要求等
            5. 时间要求：投标截止时间、有效期等
            6. 其他废标条款：特殊要求和限制条件
            
        注意：
            - 分析结果仅供参考，实际应用需人工复核
            - 分析准确性依赖于文档内容的完整性和清晰度
        """
        # 构建详细的分析提示词
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
            # 调用Qwen API进行分析（保持原有默认行为不变）
            response = self._call_qwen_api(prompt)
            # 解析并返回结构化结果
            return self._parse_tender_response(response)
        except Exception as e:
            # 异常处理，返回错误信息和默认结构
            return {
                "summary": f"分析过程中出现错误：{str(e)}",
                "invalid_items": [],
                "suggestions": ["请检查文件内容或重新尝试分析"]
            }
    
    def analyze_bid_document(self, content: str, tender_analysis: Optional[Dict] = None) -> Dict:
        """
        分析投标文件，检查合规性和废标风险
        ==================================
        
        基于招标文件要求（如果提供）对投标文件进行合规性检查，
        识别潜在的废标风险并提供改进建议。
        
        Args:
            content (str): 投标文件的文本内容
            tender_analysis (Optional[Dict]): 招标文件分析结果，可选参数
                如果提供，将进行更精确的对比分析
                
        Returns:
            Dict: 合规性检查结果字典，包含以下字段：
                - summary (str): 投标文件合规性分析摘要
                - compliance_check (Dict): 合规性检查结果
                    - overall_status (str): 整体合规状态
                    - risk_level (str): 风险等级(低/中/高)
                    - score (int): 合规性得分(0-100)
                - issues (List[Dict]): 发现的问题列表
                    - category (str): 问题类别
                    - description (str): 问题描述
                    - severity (str): 严重程度
                    - suggestion (str): 改进建议
                    - location (str): 问题位置
                - recommendations (List[str]): 总体改进建议
                
        检查维度：
            1. 资质要求满足情况
            2. 技术规格匹配程度
            3. 商务条件符合性
            4. 文件格式正确性
            5. 必要信息完整性
            6. 签名盖章完整性
            
        分析逻辑：
            - 如果提供招标文件分析结果，会进行精确对比
            - 无招标文件分析时，基于常规投标要求检查
            - 提供风险等级评估和具体改进建议
        """
        # 构建基础分析提示词
        base_prompt = f"""
        请分析以下投标文件内容，检查是否存在可能导致废标的问题。
        
        投标文件内容：
        {content}
        """
        
        # 如果提供了招标文件分析结果，增加对比分析内容
        if tender_analysis and 'invalid_items' in tender_analysis:
            # 提取招标文件中的废标条款
            invalid_items_text = "\n".join([
                f"- {item.get('category', '未知类别')}: {item.get('description', '')}"
                for item in tender_analysis['invalid_items']
            ])
            
            base_prompt += f"""
            
            招标文件中发现的废标条款：
            {invalid_items_text}
            
            请特别关注投标文件是否满足上述要求。
            """
        
        # 完整的分析提示词
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
            # 调用Qwen API进行分析（保持原有默认行为不变）
            response = self._call_qwen_api(prompt)
            # 解析并返回结构化结果
            return self._parse_bid_response(response)
        except Exception as e:
            # 异常处理，返回错误信息和默认结构
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
        调用Qwen API的核心方法
        ======================
        
        负责与阿里云百炼平台的Qwen模型进行通信，
        发送分析提示词并获取AI响应。
        
        Args:
            prompt (str): 发送给AI的提示词内容
            
        Returns:
            str: AI模型的响应文本
            
        配置参数：
            - model: 使用的AI模型版本
            - temperature: 控制响应随机性(0.3较低，保证一致性)
            - stream: 是否使用流式响应(False，获取完整响应)
            
        异常处理：
            - 网络连接错误
            - API密钥无效
            - 模型服务不可用
            - 响应格式错误
            
        注意：
            - 降低temperature以提高分析结果的一致性
            - 非流式调用确保获得完整的结构化响应
        """
        # 构建符合OpenAI格式的消息列表
        messages: List[Dict[str, Any]] = [{"role": "user", "content": prompt}]
        # 调用阿里云百炼API（作类型转换，满足SDK类型定义）
        completion = self.client.chat.completions.create(
            model=self.model,
            messages=cast(Any, messages),
            stream=False,
            temperature=0.3,
        )
        
        # 提取并返回响应内容
        return completion.choices[0].message.content or ""

    def _call_doubao_api(self, prompt: str) -> str:
        """
        调用豆包（字节方舟 Ark）API的核心方法
        ======================================

        与方舟兼容的OpenAI SDK调用方式一致，仅更换base_url与API Key。
        不改变上游提示词与下游解析逻辑。

        环境变量：
            - ARK_API_KEY: 方舟平台API Key
            - ARK_BASE_URL: 可选，自定义方舟API地址（默认华北）
            - DOUBAO_MODEL_ID: 可选，推理接入点ID或模型名
        """
        ark_api_key = os.getenv("ARK_API_KEY")
        if not ark_api_key:
            raise ValueError("缺少方舟平台API密钥，请设置环境变量 ARK_API_KEY")

        ark_client = OpenAI(
            base_url=self.ark_base_url,
            api_key=ark_api_key,
        )

        # 与Qwen一致，使用纯文本对话
        messages: List[Dict[str, Any]] = [{"role": "user", "content": prompt}]
        completion = ark_client.chat.completions.create(
            model=self.doubao_model_id,
            messages=cast(Any, messages),
            stream=False,
            temperature=0.3,
        )
        # Ark返回结构与OpenAI兼容
        return completion.choices[0].message.content or ""

    def _call_model_api(self, provider: str, prompt: str) -> str:
        """
        模型路由中间层
        ==============

        根据provider选择调用Qwen或豆包API。上游提示词与下游解析逻辑不变。

        Args:
            provider (str): "qwen" 或 "doubao"（别名：ali/dashscope、ark/volc）
            prompt (str): 传入的提示词（保持不变）

        Returns:
            str: 模型原始响应文本
        """
        key = (provider or "qwen").lower()
        if key in ("qwen", "ali", "dashscope"):
            return self._call_qwen_api(prompt)
        if key in ("doubao", "ark", "volc", "volcengine"):
            return self._call_doubao_api(prompt)
        # 未知提供方时回退到Qwen
        return self._call_qwen_api(prompt)

    # 新增：带模型选择的分析方法（保持原有方法不变，便于后续选择）
    def analyze_tender_document_with_model(self, content: str, provider: str = "qwen") -> Dict:
        """
        分析招标文件（可选择模型提供方）
        保持提示词与解析逻辑一致，仅切换底层模型调用。
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
            response = self._call_model_api(provider, prompt)
            return self._parse_tender_response(response)
        except Exception as e:
            return {
                "summary": f"分析过程中出现错误：{str(e)}",
                "invalid_items": [],
                "suggestions": ["请检查文件内容或重新尝试分析"],
            }

    def analyze_bid_document_with_model(self, content: str, tender_analysis: Optional[Dict] = None, provider: str = "qwen") -> Dict:
        """
        分析投标文件（可选择模型提供方）
        保持提示词与解析逻辑一致，仅切换底层模型调用。
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
            response = self._call_model_api(provider, prompt)
            return self._parse_bid_response(response)
        except Exception as e:
            return {
                "summary": f"分析过程中出现错误：{str(e)}",
                "compliance_check": {
                    "overall_status": "未知",
                    "risk_level": "未知",
                    "score": 0,
                },
                "issues": [],
                "recommendations": ["请检查文件内容或重新尝试分析"],
            }
    
    def _parse_tender_response(self, response: str) -> Dict:
        """
        解析招标文件分析的AI响应
        ========================
        
        将AI返回的文本响应解析为结构化的Python字典。
        支持多种解析策略以提高解析成功率。
        
        Args:
            response (str): AI模型的原始响应文本
            
        Returns:
            Dict: 解析后的招标文件分析结果
            
        解析策略：
            1. 直接JSON解析 - 响应格式完全正确时
            2. 正则提取JSON - 响应包含多余文本时
            3. 文本解析兜底 - JSON解析失败时的备选方案
            
        容错机制：
            - 自动处理格式不规范的JSON
            - 提取关键信息构建默认结构
            - 确保返回结果符合预期格式
        """
        # 优先尝试严格JSON解析
        try:
            parsed = json.loads(response)
            return self._normalize_tender_result(parsed)
        except json.JSONDecodeError:
            pass

        # 尝试从所有 ```json 代码块中提取，优先取最后一个
        for fenced in reversed(self._extract_all_fenced_json(response)):
            try:
                parsed = json.loads(fenced)
                return self._normalize_tender_result(parsed)
            except json.JSONDecodeError:
                continue

        # 使用平衡括号提取所有JSON对象，优先取最后一个
        for balanced in reversed(self._extract_all_balanced_json_objects(response)):
            try:
                parsed = json.loads(balanced)
                return self._normalize_tender_result(parsed)
            except json.JSONDecodeError:
                continue

        # 兜底：文本启发式提取（避免把模板/JSON属性当作条款）
        return {
            "summary": "解析响应时出现错误，但分析已完成",
            "invalid_items": self._extract_invalid_items_from_text(response),
            "suggestions": self._extract_suggestions_from_text(response),
        }
    
    def _parse_bid_response(self, response: str) -> Dict:
        """
        解析投标文件分析的AI响应
        ========================
        
        将AI返回的投标文件合规性检查响应解析为结构化数据。
        采用多层次解析策略确保解析成功。
        
        Args:
            response (str): AI模型的原始响应文本
            
        Returns:
            Dict: 解析后的投标文件分析结果
            
        解析策略：
            1. 直接JSON解析
            2. 正则提取JSON片段
            3. 文本模式识别兜底
            
        默认值设计：
            - 合规状态：需要人工审核
            - 风险等级：中等风险
            - 得分：50分（中等水平）
        """
        try:
            parsed = json.loads(response)
            return self._normalize_bid_result(parsed)
        except json.JSONDecodeError:
            pass

        for fenced in reversed(self._extract_all_fenced_json(response)):
            try:
                parsed = json.loads(fenced)
                return self._normalize_bid_result(parsed)
            except json.JSONDecodeError:
                continue

        for balanced in reversed(self._extract_all_balanced_json_objects(response)):
            try:
                parsed = json.loads(balanced)
                return self._normalize_bid_result(parsed)
            except json.JSONDecodeError:
                continue

        return {
            "summary": "解析响应时出现错误，但分析已完成",
            "compliance_check": {
                "overall_status": "需要人工审核",
                "risk_level": "中",
                "score": 50,
            },
            "issues": self._extract_issues_from_text(response),
            "recommendations": self._extract_suggestions_from_text(response),
        }

    def _extract_json_from_fence(self, text: str) -> Optional[str]:
        """
        保留向后兼容：返回首个代码块。
        """
        blocks = self._extract_all_fenced_json(text)
        return blocks[0] if blocks else None

    def _extract_all_fenced_json(self, text: str) -> List[str]:
        """
        提取所有 ```json ... ``` 或 ``` ... ``` 代码块内部文本，按出现顺序返回列表。
        """
        results: List[str] = []
        fence_patterns = [
            r"```json\s*(.*?)```",
            r"```\s*(.*?)```",
        ]
        for pat in fence_patterns:
            for m in re.finditer(pat, text, re.DOTALL | re.IGNORECASE):
                body = m.group(1).strip()
                if body:
                    results.append(body)
        return results

    def _extract_first_balanced_json(self, text: str) -> Optional[str]:
        """
        保留向后兼容：返回首个对象。
        """
        objs = self._extract_all_balanced_json_objects(text)
        return objs[0] if objs else None

    def _extract_all_balanced_json_objects(self, text: str) -> List[str]:
        """
        扫描文本，提取所有顶层平衡的大括号JSON对象（忽略字符串中的括号）。
        按出现顺序返回列表。
        """
        results: List[str] = []
        start_idx = None
        depth = 0
        in_string = False
        escape = False
        for i, ch in enumerate(text):
            if in_string:
                if escape:
                    escape = False
                elif ch == '\\':
                    escape = True
                elif ch == '"':
                    in_string = False
                continue
            else:
                if ch == '"':
                    in_string = True
                    continue
                if ch == '{':
                    if depth == 0:
                        start_idx = i
                    depth += 1
                elif ch == '}':
                    if depth > 0:
                        depth -= 1
                        if depth == 0 and start_idx is not None:
                            results.append(text[start_idx:i+1])
                            start_idx = None
        return results

    def _normalize_tender_result(self, data: Dict) -> Dict:
        """
        归一化招标分析结果结构，确保字段类型与键存在。
        """
        result: Dict = {}
        result["summary"] = str(data.get("summary", "")).strip()

        items = data.get("invalid_items", [])
        if isinstance(items, dict):
            items = [items]
        if not isinstance(items, list):
            items = []
        norm_items: List[Dict] = []
        for it in items:
            if not isinstance(it, dict):
                continue
            norm_items.append({
                "category": str(it.get("category", "")).strip() or "未分类",
                "description": str(it.get("description", "")).strip(),
                "requirement": str(it.get("requirement", "")).strip(),
                "severity": str(it.get("severity", "")).strip() or "中",
                "keywords": it.get("keywords", []) if isinstance(it.get("keywords", []), list) else [],
            })
        result["invalid_items"] = norm_items

        suggestions = data.get("suggestions", [])
        if isinstance(suggestions, str):
            suggestions = [suggestions]
        elif isinstance(suggestions, list):
            # 若为对象列表，尽可能取text字段或转为字符串
            tmp = []
            for s in suggestions:
                if isinstance(s, str):
                    tmp.append(s)
                elif isinstance(s, dict):
                    val = s.get("text") or s.get("suggestion") or ""
                    if val:
                        tmp.append(str(val))
            if tmp:
                suggestions = tmp
        else:
            suggestions = []
        result["suggestions"] = suggestions
        return result

    def _normalize_bid_result(self, data: Dict) -> Dict:
        """
        归一化投标分析结果结构。
        """
        result: Dict = {}
        result["summary"] = str(data.get("summary", "")).strip()

        cc = data.get("compliance_check", {})
        if not isinstance(cc, dict):
            cc = {}
        result["compliance_check"] = {
            "overall_status": str(cc.get("overall_status", "需要人工审核") or "需要人工审核"),
            "risk_level": str(cc.get("risk_level", "中") or "中"),
            "score": int(cc.get("score", 50) or 50),
        }

        issues = data.get("issues", [])
        if isinstance(issues, dict):
            issues = [issues]
        if not isinstance(issues, list):
            issues = []
        norm_issues: List[Dict] = []
        for it in issues:
            if not isinstance(it, dict):
                continue
            norm_issues.append({
                "category": str(it.get("category", "") or "未分类"),
                "description": str(it.get("description", "")),
                "severity": str(it.get("severity", "") or "中"),
                "suggestion": str(it.get("suggestion", "")),
                "location": str(it.get("location", "")),
            })
        result["issues"] = norm_issues

        recs = data.get("recommendations", [])
        if isinstance(recs, str):
            recs = [recs]
        elif isinstance(recs, list):
            tmp = []
            for r in recs:
                if isinstance(r, str):
                    tmp.append(r)
                elif isinstance(r, dict):
                    val = r.get("text") or r.get("recommendation") or ""
                    if val:
                        tmp.append(str(val))
            if tmp:
                recs = tmp
        else:
            recs = []
        result["recommendations"] = recs
        return result
    
    def _extract_invalid_items_from_text(self, text: str) -> List[Dict]:
        """
        从文本中提取废标项信息（兜底解析方法）
        ======================================
        
        当JSON解析失败时，使用文本模式识别来提取废标条款信息。
        通过关键词匹配识别可能的废标相关内容。
        
        Args:
            text (str): AI响应的原始文本
            
        Returns:
            List[Dict]: 提取的废标项列表
            
        提取逻辑：
            - 搜索包含废标相关关键词的行
            - 每行构建一个废标项字典
            - 限制返回数量避免过多无效信息
            
        关键词：
            - '废标', '无效', '不符合', '要求'
            
        注意：
            - 这是兜底方案，准确性相对较低
            - 主要确保系统不会因解析失败而崩溃
        """
        items = []
        # 按行分割文本
        lines = text.split('\n')

        def is_likely_json_line(s: str) -> bool:
            s2 = s.strip()
            if not s2:
                return False
            if s2.startswith(('{', '}', ']', '[')):
                return True
            # 典型的 JSON 属性行或模板占位
            if re.match(r'^"[A-Za-z_][A-Za-z0-9_]*"\s*:', s2):
                return True
            if '招标文件分析摘要' in s2 or '投标文件合规性分析摘要' in s2 or 'summary' in s2:
                return True
            return False

        for line in lines:
            if is_likely_json_line(line):
                continue
            if any(keyword in line for keyword in ['废标', '无效', '不符合', '要求']):
                items.append({
                    "category": "从文本提取",
                    "description": line.strip(),
                    "requirement": "",
                    "severity": "中",
                    "keywords": []
                })
        return items[:10]
    
    def _extract_issues_from_text(self, text: str) -> List[Dict]:
        """
        从文本中提取问题信息（兜底解析方法）
        ==================================
        
        当JSON解析失败时，使用文本模式识别来提取投标文件问题。
        通过关键词匹配识别可能的问题描述。
        
        Args:
            text (str): AI响应的原始文本
            
        Returns:
            List[Dict]: 提取的问题列表
            
        提取逻辑：
            - 搜索包含问题相关关键词的行
            - 每行构建一个问题字典
            - 提供通用的改进建议
            
        关键词：
            - '问题', '缺少', '不完整', '错误'
            
        默认配置：
            - 严重程度：中等
            - 改进建议：通用建议
            - 位置：未定位
        """
        issues = []
        lines = text.split('\n')

        def is_likely_json_line(s: str) -> bool:
            s2 = s.strip()
            if not s2:
                return False
            if s2.startswith(('{', '}', ']', '[')):
                return True
            if re.match(r'^"[A-Za-z_][A-Za-z0-9_]*"\s*:', s2):
                return True
            if 'summary' in s2:
                return True
            return False

        for line in lines:
            if is_likely_json_line(line):
                continue
            if any(keyword in line for keyword in ['问题', '缺少', '不完整', '错误']):
                issues.append({
                    "category": "从文本提取",
                    "description": line.strip(),
                    "severity": "中",
                    "suggestion": "请核实相关内容",
                    "location": "未定位"
                })
        return issues[:10]
    
    def _extract_suggestions_from_text(self, text: str) -> List[str]:
        """
        从文本中提取建议信息（兜底解析方法）
        ==================================
        
        当JSON解析失败时，使用文本模式识别来提取改进建议。
        通过关键词匹配识别可能的建议内容。
        
        Args:
            text (str): AI响应的原始文本
            
        Returns:
            List[str]: 提取的建议列表
            
        提取逻辑：
            - 搜索包含建议相关关键词的行
            - 直接将整行作为建议内容
            - 过滤空白和重复内容
            
        关键词：
            - '建议', '应该', '需要', '注意'
            
        数量限制：
            - 最多返回5条建议，确保信息精炼
        """
        suggestions = []
        # 按行分割文本
        lines = text.split('\n')
        
        # 遍历每行，查找包含建议相关关键词的内容
        for line in lines:
            if any(keyword in line for keyword in ['建议', '应该', '需要', '注意']):
                # 清理空白字符并添加到建议列表
                suggestion = line.strip()
                if suggestion:  # 确保不是空行
                    suggestions.append(suggestion)
        
        # 限制返回数量，保持建议精炼
        return suggestions[:5]
