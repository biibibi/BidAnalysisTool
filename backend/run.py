#!/usr/bin/env python3
"""
投标文件合规性检查工具 - 后端服务启动脚本
==========================================

本脚本是整个系统的启动入口，负责环境检查、配置验证和服务启动。
确保所有必要的配置都已正确设置后再启动Flask应用。

主要功能：
    1. 加载环境变量配置
    2. 验证API密钥配置
    3. 启动Flask Web服务
    4. 提供用户友好的错误提示

环境要求：
    - Python 3.8+
    - 已安装requirements.txt中的依赖
    - 正确配置.env文件中的API密钥

配置检查：
    - DASHSCOPE_API_KEY: 阿里云百炼API密钥
    - HOST: 服务监听地址（默认0.0.0.0）
    - PORT: 服务端口（默认5000）
    - DEBUG: 调试模式（默认True）

使用方法：
    python run.py

部署建议：
    - 开发环境：直接运行此脚本
    - 生产环境：使用Gunicorn等WSGI服务器
    - 容器部署：配合Dockerfile使用

作者：BidAnalysis Team
创建时间：2025年
版本：1.0
"""

import os
import sys
from dotenv import load_dotenv

# 加载环境变量配置文件
# 从.env文件中读取配置，如API密钥、服务参数等
load_dotenv()

# === API密钥验证 ===
# 检查阿里云百炼API密钥是否已正确配置
api_key = os.getenv("DASHSCOPE_API_KEY")
if not api_key or api_key == "your_api_key_here":
    print("❌ 错误: 请在 .env 文件中设置您的 DASHSCOPE_API_KEY")
    print("📖 获取API密钥: https://dashscope.console.aliyun.com/")
    print("💡 配置说明: 复制.env.template为.env并填入真实的API密钥")
    sys.exit(1)

# === 应用启动 ===
if __name__ == "__main__":
    """
    应用启动主函数
    ==============
    
    导入Flask应用并启动开发服务器。
    从环境变量读取服务配置参数。
    
    配置参数：
        - HOST: 监听地址，0.0.0.0允许外部访问
        - PORT: 监听端口，默认5000
        - DEBUG: 调试模式，开发时启用，生产时禁用
        
    启动信息：
        - 显示服务地址和端口
        - 提供API健康检查地址
        - 确认服务启动状态
        
    注意事项：
        - 调试模式会自动重载代码变更
        - 调试模式不应在生产环境使用
        - 0.0.0.0监听所有网络接口
    """
    # 导入Flask应用实例
    from app import app
    
    # 从环境变量获取服务配置，提供合理的默认值
    host = os.getenv("HOST", "0.0.0.0")          # 监听地址
    port = int(os.getenv("PORT", 5000))          # 监听端口
    debug = os.getenv("DEBUG", "True").lower() == "true"  # 调试模式
    
    # 显示启动信息
    print("🚀 启动投标文件合规性检查工具后端服务...")
    print(f"🌐 服务地址: http://{host}:{port}")
    print(f"❤️  API文档: http://{host}:{port}/api/health")
    print("📁 前端界面: 请打开 frontend/index.html")
    print("⚡ 按 Ctrl+C 停止服务")
    
    # 启动Flask开发服务器
    app.run(host=host, port=port, debug=debug)
