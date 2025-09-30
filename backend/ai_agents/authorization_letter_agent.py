#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
授权委托书合规核验 Agent
=======================

用途：
    针对投标文件中的《授权委托书》《法定代表人身份证明》等文件，结合多模态（文字+图片）能力，
    自动核验以下信息：
        1. 授权书/身份证图片中的法定代表人姓名、身份证号是否与正文描述一致
        2. 证件是否在有效期内（如身份证、营业执照等出现有效期字段）
        3. 授权委托关系是否完整（授权人、被授权人、职务、联系方式等）
        4. 授权文件中出现的项目名称、项目编号是否与招标文件登记信息一致

输入约定：
    process(content, context)
        content: str -> 授权委托书(或身份证明) Word 文件路径（建议 docx）
        context 可包含：
            - tender_file_id: 招标文件在数据库中的 file_id（用于回溯项目编号/名称）
            - tender_analysis_id: 直接指定某条招标分析记录（可选）
            - expected_project_id / expected_project_name: 若外部已解析好可直接传入
            - provider: 指定调用模型(qwen / doubao)，缺省依据环境变量 LLM_PROVIDER

依赖：
    - multimodal_api.MultimodalDocumentAnalyzer (多模态抽取与调用)
    - database.DatabaseManager (查询项目编号、项目名称)
    - qwen_service（底层模型兼容层，可直接通过多模态接口完成，此处主要用多模态）

模型提示词（核心需求原文）：
    你是一个投标文件审查专家。请完成以下任务：
    1. 查看附带的身份证扫描件或相关证件
    2. 判断法定代表人姓名、身份证号等信息是否与图片保持一致
    3. 判断证件是否在有效期内
    4. 项目编号与项目名称是否正确。
    请以JSON格式输出分析结果

输出（data 补充字段示例）：
    {
        "raw_model_output": "原始模型输出文本",
        "parsed_result": {"解析成功的JSON"},
        "expected_project_id": "期望项目编号",
        "expected_project_name": "期望项目名称",
        "detected_project_id": "正则初步识别的项目编号",
        "detected_project_name": "正则初步识别的项目名称",
        "project_id_match": true,
        "project_name_match": false,
        "multimodal": {"text_length": 1234, "image_count": 3}
    }

说明：
    - 若模型输出 JSON 解析失败，parsed_result 为 None，但仍返回 success=True（表示流程完成），具体错误放入 parsed_error 供前端展示。
    - 轻量正则仅辅助核对。

作者：BidAnalysis Team  (2025 / v1.0)
"""
from __future__ import annotations

import json
import os
import re
import sqlite3
import logging
from typing import Any, Dict, Optional
from pathlib import Path

from .base_agent import BaseAgent

# ----------------------------- 动态导入依赖 -----------------------------
BACKEND_DIR = Path(__file__).resolve().parent.parent

# multimodal_api
try:  # 优先常规导入
    from multimodal_api import MultimodalDocumentAnalyzer  # type: ignore
except ImportError:  # 回退基于路径的加载
    import importlib.util
    mm_path = BACKEND_DIR / 'multimodal_api.py'
    spec = importlib.util.spec_from_file_location("multimodal_api", str(mm_path))
    if spec and spec.loader:  # type: ignore
        mm_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mm_module)  # type: ignore
        MultimodalDocumentAnalyzer = mm_module.MultimodalDocumentAnalyzer  # type: ignore
    else:  # pragma: no cover
        raise

# database
try:
    from database import DatabaseManager  # type: ignore
except ImportError:
    import importlib.util
    db_path = BACKEND_DIR / 'database.py'
    spec = importlib.util.spec_from_file_location("database", str(db_path))
    if spec and spec.loader:  # type: ignore
        db_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(db_module)  # type: ignore
        DatabaseManager = db_module.DatabaseManager  # type: ignore
    else:  # pragma: no cover
        DatabaseManager = None  # type: ignore


# ----------------------------- 核心 Agent -----------------------------
class AuthorizationLetterAgent(BaseAgent):
    """授权委托书/身份证明 多模态核验 Agent"""

    BASE_PROMPT = (
        "你是一个投标文件审查专家。请完成以下任务：\n"
        "1. 查看附带的身份证扫描件或相关证件\n"
        "2. 判断法定代表人姓名、身份证号等信息是否与图片保持一致\n"
        "3. 判断证件是否在有效期内\n"
        "4. 项目编号与项目名称是否正确（请对比我提供的期望值）\n"
        "请严格以 JSON 格式输出分析结果，包含以下建议字段：\n"
        "{\n  \"legal_representative\": {\n     \"name\": \"法定代表人姓名\",\n     \"id_number\": \"身份证号或统一社会信用代码(如适用)\",\n     \"validity\": {\n         \"is_valid\": true,\n         \"valid_until\": \"证件有效期(如识别到)\"\n     }\n  },\n  \"authorization\": {\n     \"delegate_name\": \"被授权人姓名(如识别)\",\n     \"delegate_id_number\": \"被授权人身份证号(如识别)\",\n     \"completeness\": \"完整/缺少关键信息\"\n  },\n  \"project_check\": {\n     \"expected_project_id\": \"我提供的项目编号\",\n     \"found_project_id\": \"文档识别到的项目编号\",\n     \"project_id_match\": true,\n     \"expected_project_name\": \"我提供的项目名称\",\n     \"found_project_name\": \"文档识别到的项目名称\",\n     \"project_name_match\": true\n  },\n  \"issues\": [ {\n       \"type\": \"字段缺失/不一致/有效期问题等\",\n       \"description\": \"问题描述\",\n       \"severity\": \"高/中/低\"\n  } ],\n  \"summary\": \"总体核验结论与建议\"\n}"
    )

    # 简单项目编号 & 名称匹配正则（与 ProjectInfoAgent 同源精简版）
    PROJECT_ID_PATTERNS = [
        r'项目编号[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)',
        r'招标编号[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)',
        r'编号[:：]\s*([A-Za-z0-9\u4e00-\u9fff\[\]（）()_\-/]+号?)'
    ]
    PROJECT_NAME_PATTERNS = [
        r'项目名称[:：]\s*([^\n\r]{1,80})',
        r'工程名称[:：]\s*([^\n\r]{1,80})'
    ]

    def __init__(self):
        super().__init__(
            name="AuthorizationLetterAgent",
            description="核验授权委托书中的人员、证件有效性以及项目编号/名称一致性（多模态）"
        )
        self.db_manager = DatabaseManager() if 'DatabaseManager' in globals() else None
        self.logger.setLevel(logging.INFO)

    # ------------------------- 主入口 -------------------------
    def process(self, content: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        try:
            if not self.validate_input(content):
                return self.create_error_result("请输入有效的授权委托书 docx 文件路径")

            file_path = Path(content).resolve()
            if not file_path.exists():
                return self.create_error_result(f"文件不存在: {file_path}")
            if file_path.suffix.lower() not in ('.docx', '.doc'):
                return self.create_error_result("仅支持 .docx / .doc 文件")

            ctx = context or {}
            expected_project_id, expected_project_name = self._resolve_expected_project(ctx)
            provider = ctx.get('provider') or os.getenv('LLM_PROVIDER', 'qwen')

            # 构建最终提示词
            final_prompt = self._build_prompt(expected_project_id, expected_project_name)

            analyzer = MultimodalDocumentAnalyzer()
            analysis = analyzer.analyze_document(str(file_path), analysis_prompt=final_prompt, provider=provider)

            text_content = analysis.get('text_content', '')
            image_count = analysis.get('image_count', 0)
            raw_model_output = analysis.get('result')

            # 解析模型 JSON
            parsed_json, parse_error = self._parse_model_json(raw_model_output)

            # 轻量正则提取
            detected_project_id = self._first_match(text_content, self.PROJECT_ID_PATTERNS)
            detected_project_name = self._first_match(text_content, self.PROJECT_NAME_PATTERNS)

            # 匹配判断（先用模型结果 project_check，再 fallback 正则）
            model_pid = None
            model_pname = None
            if isinstance(parsed_json, dict):
                pc = parsed_json.get('project_check') or {}
                if isinstance(pc, dict):
                    model_pid = pc.get('found_project_id') or pc.get('project_id')
                    model_pname = pc.get('found_project_name') or pc.get('project_name')

            effective_found_id = model_pid or detected_project_id
            effective_found_name = model_pname or detected_project_name

            project_id_match = None
            project_name_match = None
            if expected_project_id and effective_found_id:
                project_id_match = self._normalize(expected_project_id) == self._normalize(effective_found_id)
            if expected_project_name and effective_found_name:
                project_name_match = self._normalize(expected_project_name) == self._normalize(effective_found_name)

            data: Dict[str, Any] = {
                'raw_model_output': raw_model_output,
                'parsed_result': parsed_json,
                'parsed_error': parse_error,
                'expected_project_id': expected_project_id,
                'expected_project_name': expected_project_name,
                'detected_project_id': detected_project_id,
                'detected_project_name': detected_project_name,
                'effective_found_project_id': effective_found_id,
                'effective_found_project_name': effective_found_name,
                'project_id_match': project_id_match,
                'project_name_match': project_name_match,
                'multimodal': {
                    'text_length': len(text_content),
                    'image_count': image_count
                }
            }

            # success 依据多模态分析成功与否
            if not analysis.get('success'):
                return self.create_error_result(error=analysis.get('error', '多模态分析失败'), message='授权委托书多模态分析失败')

            return self.create_success_result(data, message='授权委托书多模态核验完成')

        except Exception as e:  # pragma: no cover
            self.logger.exception("授权委托书处理失败")
            return self.create_error_result(str(e))

    # ------------------------- 辅助函数 -------------------------
    def _resolve_expected_project(self, ctx: Dict[str, Any]) -> tuple[Optional[str], Optional[str]]:
        # 1. 上下文直接给出
        pid = ctx.get('expected_project_id')
        pname = ctx.get('expected_project_name')
        if pid or pname:
            return pid, pname

        # 2. 根据 tender_file_id 查询 tender_analysis 表
        tender_file_id = ctx.get('tender_file_id')
        if tender_file_id and self.db_manager:
            info = self._query_tender_project_info_by_file(tender_file_id)
            if info:
                return info.get('project_id'), info.get('project_name')

        # 3. 根据 tender_analysis_id 查询
        tender_analysis_id = ctx.get('tender_analysis_id')
        if tender_analysis_id and self.db_manager:
            info = self._query_tender_project_info_by_analysis(tender_analysis_id)
            if info:
                return info.get('project_id'), info.get('project_name')

        return None, None

    def _query_tender_project_info_by_file(self, file_id: str) -> Optional[Dict[str, Any]]:
        try:
            conn = sqlite3.connect(self.db_manager.db_path)  # type: ignore
            cur = conn.cursor()
            cur.execute("""SELECT project_id, project_name FROM tender_analysis WHERE file_id=? ORDER BY created_time DESC LIMIT 1""", (file_id,))
            row = cur.fetchone()
            conn.close()
            if row:
                return {'project_id': row[0], 'project_name': row[1]}
        except Exception as e:  # pragma: no cover
            self.logger.warning(f"查询招标项目信息失败(file): {e}")
        return None

    def _query_tender_project_info_by_analysis(self, analysis_id: str) -> Optional[Dict[str, Any]]:
        try:
            conn = sqlite3.connect(self.db_manager.db_path)  # type: ignore
            cur = conn.cursor()
            cur.execute("""SELECT project_id, project_name FROM tender_analysis WHERE id=? LIMIT 1""", (analysis_id,))
            row = cur.fetchone()
            conn.close()
            if row:
                return {'project_id': row[0], 'project_name': row[1]}
        except Exception as e:  # pragma: no cover
            self.logger.warning(f"查询招标项目信息失败(analysis): {e}")
        return None

    def _build_prompt(self, expected_project_id: Optional[str], expected_project_name: Optional[str]) -> str:
        extra = []
        if expected_project_id:
            extra.append(f"期望项目编号: {expected_project_id}")
        if expected_project_name:
            extra.append(f"期望项目名称: {expected_project_name}")
        if not extra:
            extra.append("期望项目编号: 未提供")
            extra.append("期望项目名称: 未提供")
        extra_text = "\n".join(extra)
        return f"{self.BASE_PROMPT}\n\n{extra_text}\n请在 project_check 中务必反映我提供的期望值并进行匹配判断。"

    def _parse_model_json(self, raw: Optional[str]) -> tuple[Optional[Dict[str, Any]], Optional[str]]:
        if not raw:
            return None, '模型未返回内容'
        raw_strip = raw.strip()
        # 1. 直接解析
        for candidate in self._possible_json_strings(raw_strip):
            try:
                parsed = json.loads(candidate)
                if isinstance(parsed, dict):
                    return parsed, None
            except json.JSONDecodeError:
                continue
        return None, '无法解析为有效 JSON'

    def _possible_json_strings(self, text: str):  # 生成器
        # 代码块 ```json ... ```
        fence_patterns = [r"```json\s*(.*?)```", r"```\s*(.*?)```"]
        for pat in fence_patterns:
            for m in re.finditer(pat, text, flags=re.DOTALL | re.IGNORECASE):
                body = m.group(1).strip()
                if body.startswith('{') and body.endswith('}'):
                    yield body
        # 平衡大括号提取（简单实现）
        balanced = self._extract_balanced_json_objects(text)
        for b in balanced:
            yield b
        # 原文整体最后尝试
        if text.startswith('{') and text.endswith('}'):
            yield text

    def _extract_balanced_json_objects(self, text: str):
        results = []
        depth = 0
        start = None
        in_str = False
        esc = False
        for i, ch in enumerate(text):
            if in_str:
                if esc:
                    esc = False
                elif ch == '\\':
                    esc = True
                elif ch == '"':
                    in_str = False
                continue
            else:
                if ch == '"':
                    in_str = True
                    continue
                if ch == '{':
                    if depth == 0:
                        start = i
                    depth += 1
                elif ch == '}':
                    if depth > 0:
                        depth -= 1
                        if depth == 0 and start is not None:
                            results.append(text[start:i+1])
                            start = None
        return results

    def _first_match(self, text: str, patterns) -> Optional[str]:
        for p in patterns:
            try:
                m = re.search(p, text, re.IGNORECASE)
                if m:
                    return m.group(1).strip()
            except re.error:
                continue
        return None

    def _normalize(self, s: str) -> str:
        return re.sub(r'\s+', '', s or '').upper()


# 便捷工厂函数
def create_authorization_letter_agent() -> AuthorizationLetterAgent:
    return AuthorizationLetterAgent()
