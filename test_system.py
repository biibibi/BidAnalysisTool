#!/usr/bin/env python3
"""
系统测试脚本
用于验证投标文件合规性检查工具的各个组件是否正常工作
"""

import os
import sys
import json
from pathlib import Path

# 添加backend目录到Python路径
backend_dir = Path(__file__).parent / "backend"
sys.path.insert(0, str(backend_dir))

def test_imports():
    """测试模块导入"""
    print("测试模块导入...")
    
    try:
        from qwen_service import QwenAnalysisService
        print("✓ QwenAnalysisService 导入成功")
    except Exception as e:
        print(f"✗ QwenAnalysisService 导入失败: {e}")
        return False
    
    try:
        from file_handler import FileHandler
        print("✓ FileHandler 导入成功")
    except Exception as e:
        print(f"✗ FileHandler 导入失败: {e}")
        return False
    
    try:
        from database import DatabaseManager
        print("✓ DatabaseManager 导入成功")
    except Exception as e:
        print(f"✗ DatabaseManager 导入失败: {e}")
        return False
    
    return True

def test_database():
    """测试数据库初始化"""
    print("\n测试数据库初始化...")
    
    try:
        from database import DatabaseManager
        db = DatabaseManager("test.db")
        print("✓ 数据库初始化成功")
        
        # 清理测试数据库
        if os.path.exists("test.db"):
            os.remove("test.db")
        
        return True
    except Exception as e:
        print(f"✗ 数据库初始化失败: {e}")
        return False

def test_file_handler():
    """测试文件处理器"""
    print("\n测试文件处理器...")
    
    try:
        from file_handler import FileHandler
        handler = FileHandler()
        
        # 测试文件类型检查
        assert handler.is_allowed_file("test.pdf") == True
        assert handler.is_allowed_file("test.docx") == True
        assert handler.is_allowed_file("test.txt") == False
        
        print("✓ 文件类型检查正常")
        return True
    except Exception as e:
        print(f"✗ 文件处理器测试失败: {e}")
        return False

def test_api_key():
    """测试API密钥配置"""
    print("\n测试API密钥配置...")
    
    from dotenv import load_dotenv
    load_dotenv(backend_dir / ".env")
    
    api_key = os.getenv("DASHSCOPE_API_KEY")
    if not api_key or api_key == "your_api_key_here":
        print("⚠ API密钥未配置")
        print("请在 backend/.env 文件中设置 DASHSCOPE_API_KEY")
        return False
    else:
        print("✓ API密钥已配置")
        return True

def main():
    """主测试函数"""
    print("=== 投标文件合规性检查工具 - 系统测试 ===\n")
    
    tests = [
        ("模块导入", test_imports),
        ("数据库", test_database),
        ("文件处理", test_file_handler),
        ("API密钥", test_api_key),
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        try:
            if test_func():
                passed += 1
        except Exception as e:
            print(f"✗ {test_name}测试出现异常: {e}")
    
    print(f"\n=== 测试结果 ===")
    print(f"通过: {passed}/{total}")
    
    if passed == total:
        print("🎉 所有测试通过！系统准备就绪。")
        print("\n下一步：")
        print("1. 配置API密钥（如果还没有配置）")
        print("2. 运行 backend/run.py 启动后端服务")
        print("3. 打开 frontend/index.html 使用Web界面")
    else:
        print("❌ 部分测试失败，请检查上述错误信息。")
    
    return passed == total

if __name__ == "__main__":
    main()
