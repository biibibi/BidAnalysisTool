#!/usr/bin/env python3
"""
项目信息提取和错误检查 AI Agent
===============================

AI优先的项目编号和项目名称提取及错误检查专家。
"""

import re
import json
import sys
import os
from typing import Dict, Any, Optional, List

# 计算 backend 目录路径供后续按路径导入备用
backend_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

try:
    from database import DatabaseManager
except ImportError:
    import importlib.util
    db_path = os.path.join(backend_path, 'database.py')
    spec = importlib.util.spec_from_file_location("database", db_path)
    if spec and spec.loader:
        db_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(db_module)
        DatabaseManager = db_module.DatabaseManager  # type: ignore
    else:
        DatabaseManager = None  # type: ignore

from .base_agent import BaseAgent

class ProjectInfoAgent(BaseAgent):
    """
    项目信息提取和错误检查Agent - AI优先版本
    =========================================
    
    主要特点：
    - 以AI模型为主要提取方法
    - 正则表达式作为备选补充
    - 重试机制确保成功率
    - 支持中文项目编号识别
    """
    
    def __init__(self):
        """初始化项目信息Agent"""
        super().__init__(
            name="ProjectInfoAgent",
            description="AI优先的项目编号和项目名称提取及错误检查专家"
        )

        self.db_manager = DatabaseManager() if DatabaseManager else None
        
        # 项目编号的正则表达式模式
        self.project_id_patterns = [
            r'项目编号[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)',
            r'招标编号[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)', 
            r'项目代码[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)',
            r'工程编号[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)',
            r'编号[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)',
            r'Project\s*No\.?\s*[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)',
            r'[\u4e00-\u9fff]{1,3}\[20\d{2}\]\d+号',  # 中文项目编号格式
            r'[\u4e00-\u9fff]{2,4}20\d{2}-\d+',
            r'编号[：:]\s*\(?([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/\.]+)\)?',
        ]
        
        # 项目名称的正则表达式模式
        self.project_name_patterns = [
            r'项目名称[:：]\s*([^\n\r]{1,100})',
            r'工程名称[:：]\s*([^\n\r]{1,100})',
            r'招标项目[:：]\s*([^\n\r]{1,100})',
            r'建设项目[:：]\s*([^\n\r]{1,100})',
            r'Project\s*Name\s*[:：]\s*([^\n\r]{1,100})',
            r'([^（\n\r]{10,80})（[^）]*）项目',
            r'([^（\n\r]{10,80})项目（招标编号',
            r'([^，\n\r]{10,80}系统建设[^（\n\r]*项目)',
        ]
        
        # 历史案例关键词
        self.case_keywords = [
            '案例', '业绩', '经验', '完成', '承建', '施工过', '建设过',
            '参与', '负责', '历史', '往期', '过往', '曾经', '类似项目',
            '实施过', '承担过', '主要业绩', '项目经验', '成功案例',
            '典型案例', '参考项目', '类似经验', '相关经验'
        ]
    
    def process(self, content: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """处理项目信息提取和错误检查"""
        try:
            if not self.validate_input(content):
                return self.create_error_result("输入内容无效")
            
            self.log_processing(len(content), context)
            
            doc_type = context.get("document_type", "unknown") if context else "unknown"
            
            if doc_type == "tender":
                result = self._extract_tender_info(content)
                self._persist_project_info(context, result)
                return self.create_success_result(result, "招标文件项目信息提取完成")
            elif doc_type == "bid":
                # 检查是否需要进行错误检测
                tender_project_id = context.get("tender_project_id") if context else None
                tender_project_name = context.get("tender_project_name") if context else None
                
                if tender_project_id or tender_project_name:
                    # 执行错误检测
                    result = self._check_bid_project_errors(content, tender_project_id, tender_project_name)
                    return self.create_success_result(result, "投标文件项目信息错误检查完成")
                else:
                    # 仅提取信息，不进行检测
                    result = self._auto_extract_info(content)
                    return self.create_success_result(result, "投标文件项目信息提取完成")
            else:
                result = self._auto_extract_info(content)
                return self.create_success_result(result, "项目信息自动提取完成")
                
        except Exception as e:
            self.logger.error(f"项目信息处理失败: {str(e)}")
            return self.create_error_result(str(e))
    
    def _extract_tender_info(self, content: str) -> Dict[str, Any]:
        """从招标文件中提取项目信息 - AI优先策略"""
        # 获取文档前几页内容
        front_content = self._extract_front_pages(content, pages=3, max_length=5000)
        
        # 主要方法：AI模型提取（带重试）
        ai_result = self._extract_by_ai_with_retry(front_content, "tender")
        
        # 检查AI提取结果
        ai_project_id = ai_result.get("project_id")
        ai_project_name = ai_result.get("project_name")
        ai_confidence = ai_result.get("confidence", 0.0)
        
        # 备选方案：仅在AI失败时使用正则
        regex_project_id = None
        regex_project_name = None
        
        if (not ai_project_id and not ai_project_name) or ai_confidence < 0.3:
            self.logger.info("AI提取结果不理想，启用正则表达式补充")
            regex_project_id = self._extract_by_regex(front_content, self.project_id_patterns)
            regex_project_name = self._extract_by_regex(front_content, self.project_name_patterns)
        
        # AI优先选择
        final_project_id = ai_project_id or regex_project_id
        final_project_name = ai_project_name or regex_project_name
        
        # 计算置信度
        confidence = self._calculate_ai_priority_confidence(ai_result, final_project_id, final_project_name, ai_confidence)
        
        return {
            "document_type": "tender",
            "project_id": final_project_id,
            "project_name": final_project_name,
            "extraction_methods": {
                "primary": "ai_priority",
                "ai": ai_result,
                "regex_backup": {
                    "project_id": regex_project_id,
                    "project_name": regex_project_name,
                    "used": bool(regex_project_id or regex_project_name)
                }
            },
            "content_info": {
                "analyzed_length": len(front_content),
                "is_front_pages_only": True,
                "estimated_pages_analyzed": 3
            },
            "confidence": confidence,
            "ai_priority_strategy": True
        }

    def _persist_project_info(self, context: Optional[Dict[str, Any]], result: Dict[str, Any]):
        """将提取到的项目信息同步到数据库中"""
        if not self.db_manager or not context or not result:
            return

        doc_type = (context.get("document_type") or "").lower()
        if doc_type != "tender":
            return

        file_id = context.get("file_id")
        if not file_id:
            return

        project_id = result.get("project_id")
        project_name = result.get("project_name")

        if not project_id and not project_name:
            return

        try:
            updated = self.db_manager.update_tender_project_info(file_id, project_id, project_name)
            if not updated:
                self.logger.debug(f"未找到需要更新的招标分析记录，file_id={file_id}")
        except Exception as e:
            self.logger.error(f"项目信息写入数据库失败: {str(e)}")
    
    def _extract_by_ai_with_retry(self, content: str, doc_type: str, max_retries: int = 2) -> Dict[str, Any]:
        """AI提取（带重试机制）"""
        for attempt in range(max_retries + 1):
            try:
                self.logger.info(f"开始AI提取，尝试次数: {attempt + 1}/{max_retries + 1}")
                result = self._extract_by_ai(content, doc_type)
                
                if result and (result.get("project_id") or result.get("project_name")):
                    confidence = result.get("confidence", 0.0)
                    self.logger.info(f"AI提取成功，置信度: {confidence}")
                    return result
                elif attempt < max_retries:
                    self.logger.warning(f"AI提取结果不理想，准备重试 (尝试 {attempt + 1})")
                    continue
                else:
                    self.logger.warning("AI提取在所有尝试后仍未成功")
                    return result or {}
                    
            except Exception as e:
                self.logger.error(f"AI提取尝试 {attempt + 1} 失败: {str(e)}")
                if attempt < max_retries:
                    continue
                else:
                    return {}
        
        return {}
    
    def _extract_by_ai(self, content: str, doc_type: str) -> Dict[str, Any]:
        """使用AI模型提取项目信息"""
        try:
            # 导入AI服务
            try:
                from qwen_service import QwenAnalysisService
            except ImportError:
                # 备选导入方式
                import importlib.util
                qwen_path = os.path.join(backend_path, 'qwen_service.py')
                spec = importlib.util.spec_from_file_location("qwen_service", qwen_path)
                if spec is None or spec.loader is None:
                    raise ImportError(f"无法加载模块规范: {qwen_path}")
                qwen_module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(qwen_module)
                QwenAnalysisService = qwen_module.QwenAnalysisService

            ai_service = QwenAnalysisService()
            provider = os.getenv("LLM_PROVIDER", "qwen")
            
            # 构建提示词
            if doc_type == "tender":
                prompt = self._build_tender_extract_prompt(content)
            else:
                prompt = self._build_bid_extract_prompt(content)
            
            self.logger.info(f"开始AI提取，文档类型: {doc_type}, 内容长度: {len(content)}")
            
            # 调用AI分析
            # 根据全局提供方路由调用（保持提示与解析不变）
            try:
                response = ai_service._call_model_api(provider, prompt)
            except AttributeError:
                # 向后兼容：旧版本仅支持Qwen
                response = ai_service._call_qwen_api(prompt)
            
            # 解析响应
            result = self._parse_ai_response(response)
            
            if result:
                self.logger.info(f"AI提取成功: 项目编号={result.get('project_id')}, 项目名称={result.get('project_name')}")
                return result
            else:
                self.logger.warning("AI响应解析失败，返回空结果")
                return {}
            
        except Exception as e:
            self.logger.error(f"AI提取失败: {str(e)}")
            return {}
    
    def _build_tender_extract_prompt(self, content: str) -> str:
        """构建招标文件信息提取的提示词"""
        return f"""
        你是专业的招标文件分析专家。请从以下内容中准确提取项目编号和项目名称。

        文档内容：
        {content}

        提取规则：
        1. 项目编号识别：
           - 标识：项目编号、招标编号、项目代码、工程编号
           - 格式：KQAQ2025-001、临[2025]2663号、杭[2025]001号等
           - 支持中英文、数字、括号、连字符组合
           
        2. 项目名称识别：
           - 标识：项目名称、工程名称、招标项目
           - 通常是10-80字符的描述性文字
           
        3. 重点关注文档开头的明确标识信息

        请返回JSON格式：
        {{
            "project_id": "项目编号",
            "project_name": "项目名称", 
            "confidence": 0.95,
            "extraction_source": "提取来源说明"
        }}

        注意：未找到的信息返回null，确保JSON格式正确。
        """
    
    def _build_bid_extract_prompt(self, content: str) -> str:
        """构建投标文件信息提取的提示词"""
        return f"""
        你是专业的投标文件分析专家。请提取当前投标项目信息，排除历史案例。

        文档内容：
        {content}

        提取规则：
        1. 只提取当前投标项目信息
        2. 排除包含以下关键词的项目信息：
           - 案例、业绩、经验、完成、承建、施工过
           - 参与、负责、历史、往期、过往、曾经
        3. 重点关注投标文件封面和首页信息

        请返回JSON格式：
        {{
            "project_id": "当前项目编号",
            "project_name": "当前项目名称",
            "confidence": 0.95,
            "extraction_source": "提取来源",
            "excluded_cases": ["排除的历史案例"]
        }}
        """
    
    def _parse_ai_response(self, response: str) -> Dict[str, Any]:
        """解析AI响应"""
        if not response:
            return {}
        
        try:
            # 尝试直接解析JSON
            result = json.loads(response.strip())
            if isinstance(result, dict):
                return result
        except json.JSONDecodeError:
            # 尝试提取JSON片段
            import re
            json_pattern = r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}'
            match = re.search(json_pattern, response)
            if match:
                try:
                    result = json.loads(match.group())
                    if isinstance(result, dict):
                        return result
                except json.JSONDecodeError:
                    pass
        
        return {}
    
    def _calculate_ai_priority_confidence(self, ai_result: Dict[str, Any], 
                                         final_project_id: Optional[str], 
                                         final_project_name: Optional[str],
                                         ai_confidence: float) -> float:
        """计算AI优先策略的置信度"""
        if ai_result.get("project_id") and ai_result.get("project_name"):
            base_confidence = max(0.85, ai_confidence)
        elif ai_result.get("project_id") or ai_result.get("project_name"):
            base_confidence = max(0.75, ai_confidence)
        elif final_project_id or final_project_name:
            base_confidence = 0.6  # 正则表达式补充
        else:
            base_confidence = 0.2
        
        return max(0.0, min(1.0, base_confidence))
    
    def _extract_by_regex(self, content: str, patterns: List[str]) -> Optional[str]:
        """使用正则表达式提取信息"""
        for pattern in patterns:
            try:
                match = re.search(pattern, content, re.IGNORECASE)
                if match:
                    return match.group(1).strip()
            except re.error as e:
                self.logger.warning(f"正则表达式错误: {pattern}, 错误: {str(e)}")
                continue
        return None
    
    def _extract_front_pages(self, content: str, pages: int = 3, max_length: int = 5000) -> str:
        """提取文档前几页内容"""
        if len(content) <= max_length:
            return content
        
        # 简单的页面分割逻辑
        estimated_chars_per_page = 1800
        target_length = min(pages * estimated_chars_per_page, max_length)
        
        # 尝试在句号或换行符处截断
        cutoff_content = content[:target_length]
        last_period = cutoff_content.rfind('。')
        last_newline = cutoff_content.rfind('\n')
        
        if last_period > target_length * 0.8:
            return content[:last_period + 1]
        elif last_newline > target_length * 0.8:
            return content[:last_newline]
        else:
            return cutoff_content
    
    def _auto_extract_info(self, content: str) -> Dict[str, Any]:
        """自动提取项目信息"""
        front_content = self._extract_front_pages(content, pages=3, max_length=5000)
        
        # AI优先策略
        ai_result = self._extract_by_ai_with_retry(front_content, "tender")
        
        ai_project_id = ai_result.get("project_id")
        ai_project_name = ai_result.get("project_name")
        ai_confidence = ai_result.get("confidence", 0.0)
        
        # 备选正则
        regex_project_id = None
        regex_project_name = None
        
        if (not ai_project_id and not ai_project_name) or ai_confidence < 0.3:
            regex_project_id = self._extract_by_regex(front_content, self.project_id_patterns)
            regex_project_name = self._extract_by_regex(front_content, self.project_name_patterns)
        
        final_project_id = ai_project_id or regex_project_id
        final_project_name = ai_project_name or regex_project_name
        
        confidence = self._calculate_ai_priority_confidence(ai_result, final_project_id, final_project_name, ai_confidence)
        
        return {
            "document_type": "auto_detected",
            "project_id": final_project_id,
            "project_name": final_project_name,
            "extraction_methods": {
                "primary": "ai_priority",
                "ai": ai_result,
                "regex_backup": {
                    "project_id": regex_project_id,
                    "project_name": regex_project_name,
                    "used": bool(regex_project_id or regex_project_name)
                }
            },
            "confidence": confidence,
            "ai_priority_strategy": True
        }
    
    def _check_bid_project_errors(self, content: str, tender_project_id: Optional[str] = None, 
                                 tender_project_name: Optional[str] = None) -> Dict[str, Any]:
        """
        检查投标文件中的项目信息错误 - AI+正则双重检测策略
        =======================================================
        
        使用AI优先 + 正则表达式备选的策略，检测投标文件中的项目编号和项目名称错误。
        自动排除历史案例和其他项目的干扰信息。
        
        Args:
            content (str): 投标文件内容
            tender_project_id (Optional[str]): 招标文件的项目编号
            tender_project_name (Optional[str]): 招标文件的项目名称
            
        Returns:
            Dict[str, Any]: 检测结果，包含错误列表、置信度等信息
        """
        if not tender_project_id and not tender_project_name:
            return {
                "error": "缺少招标文件的项目信息",
                "has_errors": False,
                "errors": []
            }
        
        self.logger.info(f"开始项目信息错误检测 - 目标项目编号: {tender_project_id}, 目标项目名称: {tender_project_name}")
        
        # 初始化错误列表
        errors = []
        ai_found_project_info = {}
        
        # 方法1: AI智能检测（主要方法）
        ai_errors, ai_project_info = self._detect_errors_by_ai(content, tender_project_id, tender_project_name)
        if ai_errors:
            errors.extend(ai_errors)
            self.logger.info(f"AI检测发现 {len(ai_errors)} 个错误")
        
        # 保存AI检测过程中找到的项目信息
        ai_found_project_info = ai_project_info
        
        # 方法2: 正则表达式检测（补充方法）
        regex_errors = self._detect_errors_by_regex(content, tender_project_id, tender_project_name)
        
        # 合并去重错误（避免AI和正则重复检测相同错误）
        merged_errors = self._merge_and_deduplicate_errors(errors, regex_errors)
        
        # 使用AI检测过程中找到的项目信息，避免重复AI调用
        bid_info = {
            "project_id": ai_found_project_info.get("project_id"),
            "project_name": ai_found_project_info.get("project_name"),
            "confidence": ai_found_project_info.get("confidence", 0.0)
        }
        
        result = {
            "tender_project_id": tender_project_id,
            "tender_project_name": tender_project_name,
            "has_errors": len(merged_errors) > 0,
            "error_count": len(merged_errors),
            "errors": merged_errors,
            "bid_info": bid_info,
            "detection_methods": {
                "ai_errors_count": len(ai_errors) if ai_errors else 0,
                "regex_errors_count": len(regex_errors),
                "merged_errors_count": len(merged_errors)
            }
        }
        
        self.logger.info(f"项目信息错误检测完成 - 发现 {len(merged_errors)} 个错误")
        return result
    
    def _detect_errors_by_ai(self, content: str, tender_project_id: Optional[str], 
                           tender_project_name: Optional[str]) -> tuple:
        """
        使用AI智能检测项目信息错误
        ===========================
        
        调用AI模型分析投标文件内容，识别项目编号和项目名称的错误，
        自动排除历史案例等干扰信息。
        
        Args:
            content (str): 投标文件内容
            tender_project_id (Optional[str]): 正确的项目编号
            tender_project_name (Optional[str]): 正确的项目名称
            
        Returns:
            tuple: (错误列表, 找到的项目信息)
        """
        try:
            # 构建AI检测提示词
            prompt = self._build_error_detection_prompt(content, tender_project_id, tender_project_name)
            
            # 调用AI分析
            response = self._extract_by_ai_with_retry_for_detection(prompt)
            
            if response:
                errors = response.get("errors", [])
                found_project_info = response.get("found_project_info", {})
                
                self.logger.info(f"AI检测成功，发现 {len(errors)} 个潜在错误")
                return errors, found_project_info
            else:
                self.logger.info("AI检测未发现明显错误")
                return [], {}
                
        except Exception as e:
            self.logger.error(f"AI错误检测失败: {str(e)}")
            return [], {}
    
    def _build_error_detection_prompt(self, content: str, tender_project_id: Optional[str], 
                                    tender_project_name: Optional[str]) -> str:
        """构建AI错误检测的提示词"""
        # 截取前部分内容进行分析（避免内容过长）
        analysis_content = content[:12000] if len(content) > 12000 else content
        
        prompt = f"""
你是专业的投标文件合规检查专家。请仔细分析以下投标文件内容，检测项目编号和项目名称是否与招标文件要求一致。

招标文件要求：
- 项目编号：{tender_project_id or '未提供'}
- 项目名称：{tender_project_name or '未提供'}

投标文件内容：
{analysis_content}

检测规则：
1. 查找投标文件中出现的所有项目编号，对比是否与招标文件一致
2. 查找投标文件中出现的所有项目名称，对比是否与招标文件一致
3. 排除以下情况：
   - 历史案例、业绩介绍中的其他项目信息
   - 公司介绍中的往期项目
   - 参考案例、类似项目经验
   - 包含关键词：案例、业绩、经验、完成、承建、施工过、参与、负责、历史、往期、过往、曾经
4. 重点关注：
   - 投标文件封面的项目信息
   - 投标书正文中的项目信息
   - 技术方案中的当前项目信息
   - 授权委托书中的项目信息
   - 投标函中的项目信息

请为每个发现的错误位置分别报告，不要合并相同的错误。返回JSON格式结果：
{{
    "found_project_info": {{
        "project_id": "在投标文件中找到的项目编号",
        "project_name": "在投标文件中找到的项目名称",
        "confidence": 0.95
    }},
    "errors": [
        {{
            "type": "wrong_project_id" | "wrong_project_name" | "missing_project_id" | "missing_project_name",
            "found_value": "在投标文件中发现的值",
            "correct_value": "应该使用的正确值",
            "location": "具体位置描述（如：投标文件封面第1处、授权委托书第2处）",
            "context": "错误上下文（100字内）",
            "severity": "高" | "中" | "低",
            "description": "详细错误描述",
            "confidence": 0.95
        }}
    ],
    "analysis_summary": "整体分析总结，说明发现了多少处错误",
    "excluded_cases": ["排除的历史案例项目名称"]
}}

重要提示：
- 为每个不同位置的错误分别创建一个错误记录
- 即使是相同的错误值，如果出现在不同位置，也要分别报告
- 详细说明每个错误的具体位置
- 如果没有错误，返回空的errors数组
- 为每个错误计算并返回置信度(0-1之间)，考虑因素包括：
  * 错误的明确程度（0.9-1.0：明显错误，0.7-0.9：可能错误，0.5-0.7：不确定）
  * 上下文的相关性（是否确实是当前项目信息）
  * 是否排除了历史案例干扰
"""
        return prompt
    
    def _extract_by_ai_with_retry_for_detection(self, prompt: str, max_retries: int = 2) -> Dict[str, Any]:
        """AI错误检测（带重试机制）"""
        for attempt in range(max_retries + 1):
            try:
                self.logger.info(f"开始AI错误检测，尝试次数: {attempt + 1}/{max_retries + 1}")
                
                # 导入AI服务
                try:
                    from qwen_service import QwenAnalysisService
                except ImportError:
                    import importlib.util
                    backend_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                    qwen_path = os.path.join(backend_path, 'qwen_service.py')
                    spec = importlib.util.spec_from_file_location("qwen_service", qwen_path)
                    if spec is None or spec.loader is None:
                        raise ImportError(f"无法加载模块规范: {qwen_path}")
                    qwen_module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(qwen_module)
                    QwenAnalysisService = qwen_module.QwenAnalysisService

                ai_service = QwenAnalysisService()
                provider = os.getenv("LLM_PROVIDER", "qwen")
                try:
                    response = ai_service._call_model_api(provider, prompt)
                except AttributeError:
                    response = ai_service._call_qwen_api(prompt)
                
                # 解析响应
                result = self._parse_ai_response(response)
                
                if result and isinstance(result.get("errors"), list):
                    self.logger.info(f"AI错误检测成功，发现 {len(result['errors'])} 个错误")
                    return result
                elif attempt < max_retries:
                    self.logger.warning(f"AI错误检测结果不理想，准备重试 (尝试 {attempt + 1})")
                    continue
                else:
                    self.logger.warning("AI错误检测在所有尝试后仍未成功")
                    return {"errors": []}
                    
            except Exception as e:
                self.logger.error(f"AI错误检测尝试 {attempt + 1} 失败: {str(e)}")
                if attempt < max_retries:
                    continue
                else:
                    return {"errors": []}
        
        return {"errors": []}
    
    def _detect_errors_by_regex(self, content: str, tender_project_id: Optional[str], 
                              tender_project_name: Optional[str]) -> List[Dict[str, Any]]:
        """
        使用正则表达式检测项目信息错误（备选方法）
        =========================================
        
        Args:
            content (str): 投标文件内容
            tender_project_id (Optional[str]): 正确的项目编号
            tender_project_name (Optional[str]): 正确的项目名称
            
        Returns:
            List[Dict[str, Any]]: 正则检测到的错误列表
        """
        errors = []
        
        # 1. 检查项目编号错误
        if tender_project_id:
            errors.extend(self._check_project_id_errors_regex(content, tender_project_id))
        
        # 2. 检查项目名称错误
        if tender_project_name:
            errors.extend(self._check_project_name_errors_regex(content, tender_project_name))
        
        return errors
    
    def _check_project_id_errors_regex(self, content: str, tender_project_id: str) -> List[Dict[str, Any]]:
        """使用正则表达式检查项目编号错误"""
        errors = []
        match_count = 0  # 计数器，用于区分不同位置的相同错误
        
        for pattern in self.project_id_patterns:
            try:
                matches = re.finditer(pattern, content, re.IGNORECASE)
                for match in matches:
                    found_id = match.group(1).strip()
                    match_count += 1
                    
                    # 标准化比较（去除空格、统一大小写）
                    found_id_normalized = self._normalize_project_info(found_id)
                    tender_id_normalized = self._normalize_project_info(tender_project_id)
                    
                    if found_id_normalized != tender_id_normalized:
                        context_text = self._get_context_around_match(content, match)
                        
                        # 使用改进的历史案例识别
                        if not self._is_in_historical_context(context_text, found_id):
                            # 估算位置描述
                            position = match.start()
                            estimated_section = self._estimate_document_section(content, position)
                            
                            errors.append({
                                "type": "wrong_project_id",
                                "found_value": found_id,
                                "correct_value": tender_project_id,
                                "location": f"第{match_count}处 - {estimated_section} (字符位置: {position})",
                                "context": context_text[:100] + "..." if len(context_text) > 100 else context_text,
                                "severity": "高",
                                "description": f"第{match_count}处发现错误的项目编号 '{found_id}'，应为 '{tender_project_id}'",
                                "confidence": 0.85,
                                "detection_method": "regex",
                                "match_pattern": pattern
                            })
            except re.error as e:
                self.logger.warning(f"正则表达式错误: {pattern}, 错误: {str(e)}")
                continue
        
        return errors
    
    def _check_project_name_errors_regex(self, content: str, tender_project_name: str) -> List[Dict[str, Any]]:
        """使用正则表达式检查项目名称错误"""
        errors = []
        match_count = 0  # 计数器，用于区分不同位置的相同错误
        
        for pattern in self.project_name_patterns:
            try:
                matches = re.finditer(pattern, content, re.IGNORECASE)
                for match in matches:
                    found_name = match.group(1).strip()
                    match_count += 1
                    
                    # 计算名称相似度
                    similarity = self._calculate_name_similarity(found_name, tender_project_name)
                    
                    # 如果相似度过低，可能是错误
                    if similarity < 0.7:  # 70%相似度阈值
                        context_text = self._get_context_around_match(content, match)
                        
                        # 检查是否在历史案例上下文中
                        if not self._is_in_historical_context(context_text, found_name):
                            # 估算位置描述
                            position = match.start()
                            estimated_section = self._estimate_document_section(content, position)
                            
                            severity = "高" if similarity < 0.3 else "中"
                            errors.append({
                                "type": "wrong_project_name",
                                "found_value": found_name,
                                "correct_value": tender_project_name,
                                "location": f"第{match_count}处 - {estimated_section} (字符位置: {position})",
                                "context": context_text[:100] + "..." if len(context_text) > 100 else context_text,
                                "severity": severity,
                                "description": f"第{match_count}处发现可能错误的项目名称 '{found_name}'，与正确名称相似度仅 {similarity:.1%}",
                                "confidence": 0.8,
                                "similarity": similarity,
                                "detection_method": "regex",
                                "match_pattern": pattern
                            })
            except re.error as e:
                self.logger.warning(f"项目名称正则表达式错误: {pattern}, 错误: {str(e)}")
                continue
        
        return errors
    
    def _estimate_document_section(self, content: str, position: int) -> str:
        """估算文档位置所在的章节"""
        # 简单的位置估算
        total_length = len(content)
        relative_position = position / total_length
        
        if relative_position < 0.1:
            return "文档开头/封面区域"
        elif relative_position < 0.3:
            return "投标函/授权委托书区域"
        elif relative_position < 0.6:
            return "技术方案区域"
        elif relative_position < 0.8:
            return "商务条款区域"
        else:
            return "文档末尾区域"
    
    def _merge_and_deduplicate_errors(self, ai_errors: List[Dict[str, Any]], 
                                    regex_errors: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        合并并去重AI和正则检测的错误
        ============================
        
        避免同一个错误被重复报告。
        
        Args:
            ai_errors (List[Dict[str, Any]]): AI检测的错误
            regex_errors (List[Dict[str, Any]]): 正则检测的错误
            
        Returns:
            List[Dict[str, Any]]: 合并去重后的错误列表
        """
        merged_errors = []
        seen_errors = set()
        
        # 添加AI检测的错误（优先级高）
        for error in ai_errors:
            error_key = self._generate_error_key(error)
            if error_key not in seen_errors:
                merged_errors.append(error)
                seen_errors.add(error_key)
        
        # 添加正则检测的错误（作为补充）
        for error in regex_errors:
            error_key = self._generate_error_key(error)
            if error_key not in seen_errors:
                merged_errors.append(error)
                seen_errors.add(error_key)
        
        return merged_errors
    
    def _generate_error_key(self, error: Dict[str, Any]) -> str:
        """生成错误的唯一标识符，包含位置信息以避免过度去重"""
        # 包含位置信息，使得相同错误在不同位置被视为不同错误
        location = error.get('context', '')[:50]  # 使用上下文的前50个字符作为位置标识
        return f"{error.get('type', '')}:{error.get('found_value', '')}:{error.get('correct_value', '')}:{location}"
    

    
    def _normalize_project_info(self, text: str) -> str:
        """标准化项目信息文本（用于比较）"""
        if not text:
            return ""
        
        # 去除空格、制表符、换行符
        normalized = re.sub(r'\s+', '', text)
        
        # 统一大小写
        normalized = normalized.upper()
        
        # 去除常见的括号内容（如果不是编号的一部分）
        # 保留编号格式如 [2025]、（2025）等
        
        return normalized
    
    def _calculate_name_similarity(self, name1: str, name2: str) -> float:
        """
        计算两个项目名称的相似度
        =========================
        
        使用多种方法计算相似度：
        1. 字符级别的相似度
        2. 词汇级别的相似度
        3. 关键词匹配度
        
        Args:
            name1 (str): 第一个项目名称
            name2 (str): 第二个项目名称
            
        Returns:
            float: 相似度分数（0-1之间）
        """
        if not name1 or not name2:
            return 0.0
        
        # 标准化处理
        n1 = self._normalize_text_for_comparison(name1)
        n2 = self._normalize_text_for_comparison(name2)
        
        if n1 == n2:
            return 1.0
        
        # 方法1: 字符级相似度（Levenshtein距离）
        char_similarity = self._levenshtein_similarity(n1, n2)
        
        # 方法2: 词汇级相似度
        word_similarity = self._word_similarity(n1, n2)
        
        # 方法3: 关键词匹配度
        keyword_similarity = self._keyword_similarity(n1, n2)
        
        # 加权平均
        final_similarity = (char_similarity * 0.4 + word_similarity * 0.4 + keyword_similarity * 0.2)
        
        return min(1.0, max(0.0, final_similarity))
    
    def _normalize_text_for_comparison(self, text: str) -> str:
        """标准化文本用于相似度比较"""
        if not text:
            return ""
        
        # 去除标点符号和空格
        normalized = re.sub(r'[^\u4e00-\u9fff\w]', '', text)
        
        # 转换为小写
        return normalized.lower()
    
    def _levenshtein_similarity(self, s1: str, s2: str) -> float:
        """计算Levenshtein相似度"""
        if len(s1) == 0:
            return 0.0 if len(s2) > 0 else 1.0
        if len(s2) == 0:
            return 0.0
        
        # 动态规划计算编辑距离
        matrix = [[0] * (len(s2) + 1) for _ in range(len(s1) + 1)]
        
        for i in range(len(s1) + 1):
            matrix[i][0] = i
        for j in range(len(s2) + 1):
            matrix[0][j] = j
        
        for i in range(1, len(s1) + 1):
            for j in range(1, len(s2) + 1):
                if s1[i-1] == s2[j-1]:
                    matrix[i][j] = matrix[i-1][j-1]
                else:
                    matrix[i][j] = min(
                        matrix[i-1][j] + 1,    # 删除
                        matrix[i][j-1] + 1,    # 插入
                        matrix[i-1][j-1] + 1   # 替换
                    )
        
        max_len = max(len(s1), len(s2))
        return 1.0 - (matrix[len(s1)][len(s2)] / max_len)
    
    def _word_similarity(self, s1: str, s2: str) -> float:
        """计算词汇级相似度"""
        # 简单的词汇分割（按3-4个字符为一个词）
        words1 = set([s1[i:i+3] for i in range(len(s1)-2)] + [s1[i:i+4] for i in range(len(s1)-3)])
        words2 = set([s2[i:i+3] for i in range(len(s2)-2)] + [s2[i:i+4] for i in range(len(s2)-3)])
        
        if not words1 and not words2:
            return 1.0
        if not words1 or not words2:
            return 0.0
        
        intersection = len(words1.intersection(words2))
        union = len(words1.union(words2))
        
        return intersection / union if union > 0 else 0.0
    
    def _keyword_similarity(self, s1: str, s2: str) -> float:
        """计算关键词相似度"""
        # 项目名称中的常见关键词
        keywords = ['建设', '系统', '平台', '项目', '工程', '网络', '信息', '管理', '服务', '智能', '数字']
        
        kw1 = set([kw for kw in keywords if kw in s1])
        kw2 = set([kw for kw in keywords if kw in s2])
        
        if not kw1 and not kw2:
            return 1.0  # 如果都没有关键词，认为相似
        if not kw1 or not kw2:
            return 0.5  # 如果一个有关键词另一个没有，中等相似度
        
        intersection = len(kw1.intersection(kw2))
        union = len(kw1.union(kw2))
        
        return intersection / union if union > 0 else 0.0
    
    def _is_in_historical_context(self, context: str, project_info: str) -> bool:
        """
        改进的历史案例识别算法
        =======================
        
        使用更精确的上下文分析判断项目信息是否属于历史案例。
        
        Args:
            context (str): 上下文文本
            project_info (str): 项目信息
            
        Returns:
            bool: 是否在历史案例上下文中
        """
        if not context:
            return False
        
        context_lower = context.lower()
        
        # 强指示词（高权重）
        strong_indicators = [
            '案例', '业绩', '经验', '完成', '承建', '施工过', '建设过',
            '参与', '负责', '历史', '往期', '过往', '曾经', '类似项目',
            '实施过', '承担过', '主要业绩', '项目经验', '成功案例'
        ]
        
        # 弱指示词（低权重）
        weak_indicators = ['项目', '工程', '建设', '开发']
        
        # 排除词（如果存在则不是历史案例）
        exclusion_words = ['本项目', '此项目', '当前项目', '本次', '此次']
        
        # 检查排除词
        for exclusion in exclusion_words:
            if exclusion in context_lower:
                return False
        
        # 计算指示词权重
        strong_count = sum(1 for indicator in strong_indicators if indicator in context_lower)
        weak_count = sum(1 for indicator in weak_indicators if indicator in context_lower)
        
        # 权重计算
        total_weight = strong_count * 3 + weak_count * 1
        
        # 上下文长度调整
        context_length_factor = min(1.0, len(context) / 200)
        
        # 最终判断
        threshold = 2.0  # 调整阈值
        final_score = total_weight * context_length_factor
        
        is_historical = final_score >= threshold
        
        if is_historical:
            self.logger.debug(f"识别为历史案例: '{project_info[:30]}...', 权重: {final_score:.2f}")
        
        return is_historical
    
    def _get_context_around_match(self, content: str, match, context_length: int = 200) -> str:
        """获取匹配位置周围的上下文"""
        start = max(0, match.start() - context_length)
        end = min(len(content), match.end() + context_length)
        return content[start:end]
