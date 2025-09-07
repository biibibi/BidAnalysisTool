#!/usr/bin/env python3
"""
简单的Qwen API测试脚本
验证API密钥是否可以正常工作
"""

import os
import sys
from pathlib import Path

# 添加backend目录到Python路径
backend_dir = Path(__file__).parent / "backend"
sys.path.insert(0, str(backend_dir))

from dotenv import load_dotenv

def test_qwen_api():
    """测试Qwen API连接"""
    print("测试Qwen API连接...")
    
    # 加载环境变量
    load_dotenv(backend_dir / ".env")
    
    api_key = os.getenv("DASHSCOPE_API_KEY")
    if not api_key or api_key == "your_api_key_here":
        print("❌ API密钥未配置")
        print("请在 backend/.env 文件中设置您的阿里云百炼API密钥")
        return False
    
    try:
        from qwen_service import QwenAnalysisService
        service = QwenAnalysisService()
        
        # 测试简单的分析
        test_content = """
        招标公告
        
        1. 投标人资质要求：
        - 具有有效的营业执照
        - 具有建筑工程施工总承包三级及以上资质
        - 近三年无重大质量事故
        
        2. 技术要求：
        - 符合国家相关技术标准
        - 产品质量保证期不少于2年
        
        3. 商务要求：
        - 投标有效期：90天
        - 交货期：合同签订后30天内
        """
        
        print("正在调用Qwen API进行测试分析...")
        result = service.analyze_tender_document(test_content)
        
        print("✓ API调用成功！")
        print(f"分析摘要：{result.get('summary', '无摘要')}")
        print(f"发现废标项数量：{len(result.get('invalid_items', []))}")
        
        return True
        
    except Exception as e:
        print(f"❌ API测试失败：{e}")
        return False

if __name__ == "__main__":
    print("=== Qwen API 测试 ===\n")
    
    if test_qwen_api():
        print("\n🎉 API测试成功！可以开始使用系统了。")
    else:
        print("\n❌ API测试失败，请检查配置。")
        
    print("\n使用说明：")
    print("1. 运行 backend/run.py 启动后端服务")
    print("2. 打开 frontend/index.html 使用Web界面")
