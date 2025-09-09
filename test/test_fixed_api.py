#!/usr/bin/env python3
"""
测试修复后的API调用
"""

import os
import sys
from pathlib import Path

# 添加backend目录到Python路径
backend_dir = Path(__file__).parent / "backend"
sys.path.insert(0, str(backend_dir))

from dotenv import load_dotenv

def test_fixed_api():
    """测试修复后的API调用"""
    print("测试修复后的Qwen API调用...")
    
    # 加载环境变量
    load_dotenv(backend_dir / ".env")
    
    try:
        from qwen_service import QwenAnalysisService
        service = QwenAnalysisService()
        
        # 测试招标文件分析
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
        - 投标文件未按要求装订的，按废标处理
        """
        
        print("正在分析招标文件...")
        result = service.analyze_tender_document(test_content)
        
        print("✓ 招标文件分析成功！")
        print(f"分析摘要：{result.get('summary', '无摘要')}")
        print(f"发现废标项数量：{len(result.get('invalid_items', []))}")
        
        # 显示详细结果
        if result.get('invalid_items'):
            print("\n发现的废标条款：")
            for i, item in enumerate(result['invalid_items'][:3], 1):  # 只显示前3个
                print(f"{i}. {item.get('description', '无描述')}")
        
        # 测试投标文件分析
        bid_content = """
        投标文件
        
        我公司具有有效营业执照，建筑工程施工总承包二级资质。
        近三年无重大质量事故。
        
        技术方案符合国家相关标准，产品质量保证期为3年。
        
        投标有效期：90天
        交货期：合同签订后25天内完成交货
        """
        
        print("\n正在分析投标文件...")
        bid_result = service.analyze_bid_document(bid_content, result)
        
        print("✓ 投标文件分析成功！")
        print(f"合规性状态：{bid_result.get('compliance_check', {}).get('overall_status', '未知')}")
        print(f"风险等级：{bid_result.get('compliance_check', {}).get('risk_level', '未知')}")
        
        return True
        
    except Exception as e:
        print(f"❌ API测试失败：{e}")
        return False

if __name__ == "__main__":
    print("=== 修复后的API测试 ===\n")
    
    if test_fixed_api():
        print("\n🎉 API修复成功！系统可以正常使用了。")
        print("\n使用说明：")
        print("1. 后端服务已在运行")
        print("2. 打开浏览器访问 frontend/index.html")
        print("3. 上传文件并进行分析")
    else:
        print("\n❌ API仍有问题，请检查错误信息。")
