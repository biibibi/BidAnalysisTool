#!/usr/bin/env python3
"""
文档项目信息分析脚本
==================

用于分析招标文件和投标文件中的项目名称及编号，
检查投标文件中是否存在项目信息错误。
"""

import os
import sys
import json
import traceback

# 添加project root path和backend路径
project_root = os.path.dirname(os.path.dirname(__file__))
backend_path = os.path.join(project_root, 'backend')
if project_root not in sys.path:
    sys.path.insert(0, project_root)
if backend_path not in sys.path:
    sys.path.insert(0, backend_path)

# 现在可以正确导入
from backend.file_handler import FileHandler
from backend.ai_agents.project_info_agent import ProjectInfoAgent

def main():
    """主函数：分析招标文件和投标文件的项目信息"""
    print("=" * 60)
    print("招标文件与投标文件项目信息对比分析")
    print("=" * 60)
    
    # 文件路径
    tender_file = r".\testfile\招标文件.docx"
    bid_file = r".\testfile\投标文件.docx"

    # 检查文件是否存在
    if not os.path.exists(tender_file):
        print(f"❌ 招标文件不存在: {tender_file}")
        return
    if not os.path.exists(bid_file):
        print(f"❌ 投标文件不存在: {bid_file}")
        return
    
    print(f"📄 招标文件: {os.path.basename(tender_file)}")
    print(f"📄 投标文件: {os.path.basename(bid_file)}")
    print()
    
    # 初始化文件处理器和AI代理
    file_handler = FileHandler()
    project_agent = ProjectInfoAgent()
    
    try:
        # 1. 提取招标文件内容
        print("🔍 正在提取招标文件内容...")
        tender_content = file_handler.extract_content(tender_file)
        print(f"✅ 招标文件内容提取完成，字符数: {len(tender_content)}")
        
        # 2. 从招标文件中提取项目信息
        print("\n🤖 正在分析招标文件项目信息...")
        tender_result = project_agent.process(
            tender_content, 
            context={"document_type": "tender"}
        )
        
        if not tender_result.get("success"):
            print(f"❌ 招标文件分析失败: {tender_result.get('error', '未知错误')}")
            return
        
        tender_data = tender_result.get("data", {})  # 修复：使用 "data" 而不是 "result"
        tender_project_id = tender_data.get("project_id")
        tender_project_name = tender_data.get("project_name")
        tender_confidence = tender_data.get("confidence", 0)
        
        # 调试信息
        print(f"\n🔧 调试信息:")
        print(f"   complete result: {json.dumps(tender_result, ensure_ascii=False, indent=2)[:500]}...")
        
        print("\n📊 招标文件项目信息:")
        print(f"   项目编号: {tender_project_id or '未找到'}")
        print(f"   项目名称: {tender_project_name or '未找到'}")
        print(f"   提取置信度: {tender_confidence:.2f}")
        
        # 尝试从AI结果中直接获取
        ai_methods = tender_data.get("extraction_methods", {})
        ai_result = ai_methods.get("ai", {})
        if ai_result:
            print(f"\n🔧 AI提取原始结果:")
            print(f"   AI项目编号: {ai_result.get('project_id')}")
            print(f"   AI项目名称: {ai_result.get('project_name')}")
            print(f"   AI置信度: {ai_result.get('confidence')}")
        
        if not tender_project_id and not tender_project_name:
            print("⚠️  招标文件中未找到项目信息，无法进行对比分析")
            return
        
        # 3. 提取投标文件内容
        print("\n🔍 正在提取投标文件内容...")
        bid_content = file_handler.extract_content(bid_file)
        print(f"✅ 投标文件内容提取完成，字符数: {len(bid_content)}")
        
        # 4. 检查投标文件中的项目信息错误
        print("\n🔎 正在检查投标文件项目信息错误...")
        bid_check_result = project_agent.process(
            bid_content,
            context={
                "document_type": "bid",
                "tender_project_id": tender_project_id,
                "tender_project_name": tender_project_name
            }
        )
        
        if not bid_check_result.get("success"):
            print(f"❌ 投标文件检查失败: {bid_check_result.get('error', '未知错误')}")
            return
        
        bid_data = bid_check_result.get("data", {})  # 修复：使用 "data" 而不是 "result"
        has_errors = bid_data.get("has_errors", False)
        errors = bid_data.get("errors", [])
        error_count = bid_data.get("error_count", 0)
        
        # 5. 显示分析结果
        print("\n" + "=" * 60)
        print("分析结果")
        print("=" * 60)
        
        print(f"📋 招标文件项目信息:")
        print(f"   项目编号: {tender_project_id or '未识别'}")
        print(f"   项目名称: {tender_project_name or '未识别'}")
        
        print(f"\n🔍 投标文件错误检查结果:")
        if has_errors:
            print(f"   ❌ 发现 {error_count} 个错误")
            print("\n   错误详情:")
            for i, error in enumerate(errors, 1):
                print(f"   {i}. {error.get('description', '未知错误')}")
                print(f"      错误类型: {error.get('type', '未知')}")
                print(f"      严重程度: {error.get('severity', '未知')}")
                print(f"      发现的值: {error.get('found_value', '未知')}")
                print(f"      正确的值: {error.get('correct_value', '未知')}")
                print(f"      上下文: {error.get('context', '无')[:100]}...")
                print()
        else:
            print("   ✅ 未发现项目信息错误")
        
        # 6. 保存详细结果到文件
        result_file = "project_analysis_result.json"
        full_result = {
            "analysis_time": tender_result.get("timestamp"),
            "files": {
                "tender_file": os.path.basename(tender_file),
                "bid_file": os.path.basename(bid_file)
            },
            "tender_analysis": tender_data,
            "bid_analysis": bid_data,
            "summary": {
                "tender_project_id": tender_project_id,
                "tender_project_name": tender_project_name,
                "has_errors": has_errors,
                "error_count": error_count,
                "errors": errors
            }
        }
        
        with open(result_file, 'w', encoding='utf-8') as f:
            json.dump(full_result, f, ensure_ascii=False, indent=2)
        
        print(f"\n📄 详细分析结果已保存到: {result_file}")
        
    except Exception as e:
        print(f"\n❌ 分析过程中发生错误: {str(e)}")
        print("\n详细错误信息:")
        traceback.print_exc()

if __name__ == "__main__":
    main()
