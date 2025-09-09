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

# 添加backend目录到路径以解决导入问题
backend_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if backend_path not in sys.path:
    sys.path.insert(0, backend_path)

from ai_agents.base_agent import BaseAgent

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
                return self.create_success_result(result, "招标文件项目信息提取完成")
            elif doc_type == "bid":
                tender_project_id = context.get("tender_project_id") if context else None
                tender_project_name = context.get("tender_project_name") if context else None
                result = self._check_bid_project_errors(content, tender_project_id, tender_project_name)
                return self.create_success_result(result, "投标文件项目信息错误检查完成")
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
                qwen_module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(qwen_module)
                QwenAnalysisService = qwen_module.QwenAnalysisService
            
            ai_service = QwenAnalysisService()
            
            # 构建提示词
            if doc_type == "tender":
                prompt = self._build_tender_extract_prompt(content)
            else:
                prompt = self._build_bid_extract_prompt(content)
            
            self.logger.info(f"开始AI提取，文档类型: {doc_type}, 内容长度: {len(content)}")
            
            # 调用AI分析
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
        """检查投标文件中的项目信息错误"""
        if not tender_project_id and not tender_project_name:
            return {
                "error": "缺少招标文件的项目信息",
                "has_errors": False,
                "errors": []
            }
        
        # 简化的错误检查（仅使用正则表达式）
        errors = []
        
        # 检查项目编号错误
        if tender_project_id:
            for pattern in self.project_id_patterns:
                try:
                    matches = re.finditer(pattern, content, re.IGNORECASE)
                    for match in matches:
                        found_id = match.group(1).strip()
                        if found_id.upper() != tender_project_id.upper():
                            context_text = self._get_context_around_match(content, match)
                            if not self._is_in_case_context(context_text):
                                errors.append({
                                    "type": "wrong_project_id",
                                    "found_value": found_id,
                                    "correct_value": tender_project_id,
                                    "location": f"位置: {match.start()}-{match.end()}",
                                    "context": context_text[:200] + "..." if len(context_text) > 200 else context_text,
                                    "severity": "高",
                                    "description": f"发现错误的项目编号 '{found_id}'，应为 '{tender_project_id}'"
                                })
                except re.error:
                    continue
        
        return {
            "tender_project_id": tender_project_id,
            "tender_project_name": tender_project_name,
            "has_errors": len(errors) > 0,
            "error_count": len(errors),
            "errors": errors
        }
    
    def _get_context_around_match(self, content: str, match, context_length: int = 200) -> str:
        """获取匹配位置周围的上下文"""
        start = max(0, match.start() - context_length)
        end = min(len(content), match.end() + context_length)
        return content[start:end]
    
    def _is_in_case_context(self, context: str) -> bool:
        """判断是否在历史案例上下文中"""
        for keyword in self.case_keywords:
            if keyword in context:
                return True
        return False
