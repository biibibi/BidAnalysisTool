#!/usr/bin/env python3
"""
AI Agent 基础类
===============

定义所有AI Agent的基础接口和通用功能。
所有具体的Agent都应该继承这个基类。

基础功能：
    1. 统一的初始化接口
    2. 通用的错误处理机制
    3. 标准的结果格式
    4. 日志记录功能

设计原则：
    - 单一职责：每个Agent专注一个特定任务
    - 可扩展性：易于添加新的Agent
    - 一致性：所有Agent使用相同的接口
    - 可测试性：提供清晰的输入输出格式

作者：BidAnalysis Team
创建时间：2025年
版本：1.0
"""

from abc import ABC, abstractmethod
from typing import Dict, Any, Optional
import logging
from datetime import datetime

class BaseAgent(ABC):
    """
    AI Agent 基础抽象类
    ===================
    
    所有AI Agent的基类，定义了标准的接口和通用功能。
    
    主要功能：
        - 定义标准的处理接口
        - 提供通用的错误处理
        - 标准化结果格式
        - 日志记录支持
    
    使用方法：
        1. 继承BaseAgent类
        2. 实现process()抽象方法
        3. 可选择重写其他方法以自定义行为
    """
    
    def __init__(self, name: str, description: str = ""):
        """
        初始化Agent
        
        Args:
            name (str): Agent的名称
            description (str): Agent的描述信息
        """
        self.name = name
        self.description = description
        self.logger = logging.getLogger(f"AI_Agent.{name}")
        
        # 设置日志格式
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.INFO)
    
    @abstractmethod
    def process(self, content: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        处理输入内容的抽象方法
        
        Args:
            content (str): 要处理的文本内容
            context (Optional[Dict[str, Any]]): 可选的上下文信息
            
        Returns:
            Dict[str, Any]: 处理结果，包含标准格式的输出
            
        标准输出格式：
            {
                "success": bool,           # 处理是否成功
                "data": dict,             # 处理结果数据
                "message": str,           # 处理消息
                "error": str,             # 错误信息（如果有）
                "agent_name": str,        # Agent名称
                "timestamp": str          # 处理时间戳
            }
        """
        pass
    
    def validate_input(self, content: str) -> bool:
        """
        验证输入内容的有效性
        
        Args:
            content (str): 输入内容
            
        Returns:
            bool: 输入是否有效
        """
        if not content or not isinstance(content, str):
            self.logger.warning("输入内容为空或非字符串类型")
            return False
        
        if len(content.strip()) == 0:
            self.logger.warning("输入内容为空白字符")
            return False
            
        return True
    
    def create_success_result(self, data: Dict[str, Any], message: str = "处理成功") -> Dict[str, Any]:
        """
        创建成功结果
        
        Args:
            data (Dict[str, Any]): 结果数据
            message (str): 成功消息
            
        Returns:
            Dict[str, Any]: 标准格式的成功结果
        """
        return {
            "success": True,
            "data": data,
            "message": message,
            "error": None,
            "agent_name": self.name,
            "timestamp": datetime.now().isoformat()
        }
    
    def create_error_result(self, error: str, message: str = "处理失败") -> Dict[str, Any]:
        """
        创建错误结果
        
        Args:
            error (str): 错误信息
            message (str): 错误消息
            
        Returns:
            Dict[str, Any]: 标准格式的错误结果
        """
        return {
            "success": False,
            "data": {},
            "message": message,
            "error": error,
            "agent_name": self.name,
            "timestamp": datetime.now().isoformat()
        }
    
    def log_processing(self, content_length: int, context: Optional[Dict] = None):
        """
        记录处理日志
        
        Args:
            content_length (int): 处理内容的长度
            context (Optional[Dict]): 上下文信息
        """
        self.logger.info(f"开始处理内容，长度: {content_length} 字符")
        if context:
            self.logger.debug(f"上下文信息: {context}")
    
    def get_agent_info(self) -> Dict[str, str]:
        """
        获取Agent信息
        
        Returns:
            Dict[str, str]: Agent的基本信息
        """
        return {
            "name": self.name,
            "description": self.description,
            "type": self.__class__.__name__
        }
