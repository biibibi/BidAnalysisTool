#!/usr/bin/env python3
"""
投标文件合规性检查工具 - 后端服务启动脚本
"""

import os
import sys
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 检查API密钥
api_key = os.getenv("DASHSCOPE_API_KEY")
if not api_key or api_key == "your_api_key_here":
    print("错误: 请在 .env 文件中设置您的 DASHSCOPE_API_KEY")
    print("获取API密钥: https://dashscope.console.aliyun.com/")
    sys.exit(1)

# 启动应用
if __name__ == "__main__":
    from app import app
    
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("DEBUG", "True").lower() == "true"
    
    print(f"启动投标文件合规性检查工具后端服务...")
    print(f"服务地址: http://{host}:{port}")
    print(f"API文档: http://{host}:{port}/api/health")
    
    app.run(host=host, port=port, debug=debug)
