#!/usr/bin/env python3
"""
Word 拆分器测试脚本
用于测试和调试 word_splitter.py 中的问题
"""

import os
import sys

# 添加 backend 目录到 Python 路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backend'))

def test_word_splitter():
    """测试 Word 拆分功能"""
    
    print("🧪 Word 拆分器测试")
    print("=" * 50)
    
    # 检查测试文档是否存在
    test_files = [
        "backend/uploads/投标文件.docx",
        "test/testfile/投标文件.docx",
        "test/testfile/招标文件.docx"
    ]
    
    test_file = None
    for file_path in test_files:
        if os.path.exists(file_path):
            test_file = file_path
            break
    
    if not test_file:
        print("❌ 未找到测试用的 Word 文档")
        print("📁 搜索路径:")
        for file_path in test_files:
            print(f"   - {file_path} {'✅' if os.path.exists(file_path) else '❌'}")
        return False
    
    print(f"📄 使用测试文档: {test_file}")
    
    try:
        # 导入拆分模块
        from ai_agents.word_splitter import split_word_document_stable
        
        # 执行拆分测试
        print("\n🔄 开始拆分测试...")
        split_files = split_word_document_stable(test_file)
        
        if split_files:
            print(f"✅ 拆分成功！生成了 {len(split_files)} 个文件")
            for i, file_info in enumerate(split_files[:5], 1):  # 只显示前5个
                print(f"   {i}. {file_info['filename']} - {file_info['title']}")
            if len(split_files) > 5:
                print(f"   ... 还有 {len(split_files) - 5} 个文件")
        else:
            print("❌ 拆分失败，没有生成文件")
            return False
            
    except ImportError as e:
        print(f"❌ 导入错误: {e}")
        print("💡 请确保已安装 pywin32: pip install pywin32")
        return False
    except Exception as e:
        print(f"❌ 拆分过程出错: {e}")
        print(f"❌ 错误类型: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        return False
    
    return True

def check_dependencies():
    """检查依赖项"""
    print("\n🔍 检查依赖项...")
    
    dependencies = [
        ('win32com.client', 'pywin32'),
        ('pythoncom', 'pywin32'),
        ('pywintypes', 'pywin32')
    ]
    
    all_ok = True
    for module_name, package_name in dependencies:
        try:
            __import__(module_name)
            print(f"✅ {module_name} - 已安装")
        except ImportError:
            print(f"❌ {module_name} - 未安装，需要: pip install {package_name}")
            all_ok = False
    
    return all_ok

def check_word_availability():
    """检查 Word 是否可用"""
    print("\n🔍 检查 Microsoft Word...")
    
    try:
        import win32com.client
        import pythoncom
        
        pythoncom.CoInitialize()
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.Quit()
        pythoncom.CoUninitialize()
        
        print("✅ Microsoft Word 可用")
        return True
        
    except Exception as e:
        print(f"❌ Microsoft Word 不可用: {e}")
        print("💡 请确保已安装 Microsoft Office Word")
        return False

if __name__ == "__main__":
    print("🚀 Word 拆分器诊断工具")
    print("=" * 60)
    
    # 步骤1: 检查依赖
    if not check_dependencies():
        print("\n❌ 依赖检查失败，请先安装缺失的包")
        sys.exit(1)
    
    # 步骤2: 检查 Word
    if not check_word_availability():
        print("\n❌ Word 检查失败，请确保 Microsoft Office 已正确安装")
        sys.exit(1)
    
    # 步骤3: 测试拆分功能
    if test_word_splitter():
        print("\n🎉 所有测试通过！")
    else:
        print("\n❌ 测试失败")
        sys.exit(1)