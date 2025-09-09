#!/usr/bin/env python3
"""
AI Agent 模块包
===============

本模块包含各种专业化的AI Agent，用于处理不同的业务场景。
每个Agent都专注于特定的任务，便于维护和扩展。

当前支持的Agent：
    - ProjectInfoAgent: 项目信息提取和匹配
    - 未来可扩展: ComplianceAgent, QualificationAgent 等

架构设计：
    - BaseAgent: 所有Agent的基类
    - 具体Agent: 继承BaseAgent实现特定功能
    - AgentManager: 统一管理和调度各个Agent

作者：BidAnalysis Team
创建时间：2025年
版本：1.0
"""

from .base_agent import BaseAgent
from .project_info_agent import ProjectInfoAgent
from .agent_manager import AgentManager

__all__ = ['BaseAgent', 'ProjectInfoAgent', 'AgentManager']
