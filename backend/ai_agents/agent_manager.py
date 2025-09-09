#!/usr/bin/env python3
"""
AI Agent 管理器
===============

统一管理和调度各种AI Agent的管理器。
提供Agent的注册、发现、调用等功能。

主要功能：
    1. Agent注册和管理
    2. 统一的调用接口
    3. Agent状态监控
    4. 结果聚合和处理

设计优势：
    - 解耦：业务逻辑与具体Agent解耦
    - 扩展：易于添加新的Agent
    - 管理：统一的Agent生命周期管理
    - 监控：Agent性能和状态监控

作者：BidAnalysis Team
创建时间：2025年
版本：1.0
"""

from typing import Dict, Any, Optional, List, Type
import logging
from .base_agent import BaseAgent
from .project_info_agent import ProjectInfoAgent

class AgentManager:
    """
    AI Agent 管理器
    ===============
    
    负责管理所有AI Agent的注册、调用和监控。
    
    功能：
        - Agent注册和发现
        - 统一调用接口
        - 结果聚合
        - 错误处理
        - 性能监控
    """
    
    def __init__(self):
        """
        初始化Agent管理器
        """
        self.agents: Dict[str, BaseAgent] = {}
        self.logger = logging.getLogger("AgentManager")
        
        # 自动注册内置Agent
        self._register_builtin_agents()
    
    def _register_builtin_agents(self):
        """
        注册内置的Agent
        """
        try:
            # 注册项目信息Agent
            project_agent = ProjectInfoAgent()
            self.register_agent(project_agent)
            
            self.logger.info("内置Agent注册完成")
            
        except Exception as e:
            self.logger.error(f"内置Agent注册失败: {str(e)}")
    
    def register_agent(self, agent: BaseAgent) -> bool:
        """
        注册一个Agent
        
        Args:
            agent (BaseAgent): 要注册的Agent实例
            
        Returns:
            bool: 注册是否成功
        """
        try:
            if not isinstance(agent, BaseAgent):
                self.logger.error("Agent必须继承BaseAgent类")
                return False
            
            agent_name = agent.name
            if agent_name in self.agents:
                self.logger.warning(f"Agent '{agent_name}' 已存在，将被覆盖")
            
            self.agents[agent_name] = agent
            self.logger.info(f"Agent '{agent_name}' 注册成功")
            return True
            
        except Exception as e:
            self.logger.error(f"Agent注册失败: {str(e)}")
            return False
    
    def unregister_agent(self, agent_name: str) -> bool:
        """
        注销一个Agent
        
        Args:
            agent_name (str): Agent名称
            
        Returns:
            bool: 注销是否成功
        """
        try:
            if agent_name in self.agents:
                del self.agents[agent_name]
                self.logger.info(f"Agent '{agent_name}' 注销成功")
                return True
            else:
                self.logger.warning(f"Agent '{agent_name}' 不存在")
                return False
                
        except Exception as e:
            self.logger.error(f"Agent注销失败: {str(e)}")
            return False
    
    def get_agent(self, agent_name: str) -> Optional[BaseAgent]:
        """
        获取指定的Agent
        
        Args:
            agent_name (str): Agent名称
            
        Returns:
            Optional[BaseAgent]: Agent实例，如果不存在返回None
        """
        return self.agents.get(agent_name)
    
    def list_agents(self) -> List[Dict[str, str]]:
        """
        列出所有已注册的Agent
        
        Returns:
            List[Dict[str, str]]: Agent信息列表
        """
        return [agent.get_agent_info() for agent in self.agents.values()]
    
    def process_with_agent(self, agent_name: str, content: str, 
                          context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        使用指定Agent处理内容
        
        Args:
            agent_name (str): Agent名称
            content (str): 要处理的内容
            context (Optional[Dict[str, Any]]): 上下文信息
            
        Returns:
            Dict[str, Any]: 处理结果
        """
        try:
            agent = self.get_agent(agent_name)
            if not agent:
                return {
                    "success": False,
                    "error": f"Agent '{agent_name}' 不存在",
                    "available_agents": list(self.agents.keys())
                }
            
            # 调用Agent处理
            result = agent.process(content, context)
            
            # 添加管理器信息
            result["manager_info"] = {
                "agent_used": agent_name,
                "total_agents": len(self.agents)
            }
            
            return result
            
        except Exception as e:
            self.logger.error(f"Agent处理失败: {str(e)}")
            return {
                "success": False,
                "error": f"Agent处理失败: {str(e)}",
                "agent_name": agent_name
            }
    
    def extract_project_info(self, content: str, document_type: str = "auto") -> Dict[str, Any]:
        """
        提取项目信息的便捷方法
        
        Args:
            content (str): 文档内容
            document_type (str): 文档类型 ("tender", "bid", "auto")
            
        Returns:
            Dict[str, Any]: 项目信息提取结果
        """
        context = {"document_type": document_type}
        return self.process_with_agent("ProjectInfoAgent", content, context)
    
    def match_project_info(self, bid_content: str, tender_info: Dict[str, Any]) -> Dict[str, Any]:
        """
        匹配项目信息的便捷方法
        
        Args:
            bid_content (str): 投标文件内容
            tender_info (Dict[str, Any]): 招标文件项目信息
            
        Returns:
            Dict[str, Any]: 项目信息匹配结果
        """
        context = {
            "document_type": "bid",
            "tender_info": tender_info
        }
        return self.process_with_agent("ProjectInfoAgent", bid_content, context)
    
    def get_agent_statistics(self) -> Dict[str, Any]:
        """
        获取Agent统计信息
        
        Returns:
            Dict[str, Any]: 统计信息
        """
        return {
            "total_agents": len(self.agents),
            "agent_names": list(self.agents.keys()),
            "agent_types": [agent.__class__.__name__ for agent in self.agents.values()]
        }
    
    def health_check(self) -> Dict[str, Any]:
        """
        健康检查
        
        Returns:
            Dict[str, Any]: 健康状态
        """
        healthy_agents = []
        unhealthy_agents = []
        
        for name, agent in self.agents.items():
            try:
                # 简单的健康检查
                info = agent.get_agent_info()
                if info and info.get("name"):
                    healthy_agents.append(name)
                else:
                    unhealthy_agents.append(name)
            except Exception as e:
                unhealthy_agents.append(f"{name}: {str(e)}")
        
        return {
            "status": "healthy" if not unhealthy_agents else "unhealthy",
            "total_agents": len(self.agents),
            "healthy_agents": healthy_agents,
            "unhealthy_agents": unhealthy_agents
        }

# 创建全局Agent管理器实例
agent_manager = AgentManager()
