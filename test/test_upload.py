#!/usr/bin/env python3
"""测试招标文件上传和分析功能"""

import requests
import json

def test_tender_analysis():
    """测试招标文件上传和分析"""
    print("开始测试招标文件分析功能...")
    
    # 使用绝对路径
    file_path = r"d:\BidAnalysisTool\test\testfile\招标文件.docx"
    
    try:
        # 测试文件上传
        print("1. 测试文件上传...")
        files = {'file': open(file_path, 'rb')}
        response = requests.post('http://127.0.0.1:5000/api/upload', files=files)
        files['file'].close()
        
        print(f"   上传响应状态码: {response.status_code}")
        result = response.json()
        
        if response.status_code == 200:
            file_id = result['file_id']
            print(f"   ✅ 文件上传成功！文件ID: {file_id}")
            
            # 测试招标文件分析
            print("\n2. 测试招标文件分析...")
            analysis_response = requests.post(
                'http://127.0.0.1:5000/api/analyze/tender', 
                json={'file_id': file_id}, 
                timeout=120
            )
            
            print(f"   分析响应状态码: {analysis_response.status_code}")
            analysis_result = analysis_response.json()
            
            if analysis_response.status_code == 200:
                print("   ✅ 招标文件分析成功！")
                print(f"   分析ID: {analysis_result.get('analysis_id')}")
                
                # 显示分析结果摘要
                if 'result' in analysis_result:
                    data = analysis_result['result']
                    print("\n📊 === 分析结果摘要 ===")
                    print(f"   项目名称: {data.get('project_name', '未识别')}")
                    print(f"   招标编号: {data.get('tender_number', '未识别')}")
                    print(f"   废标条款数量: {len(data.get('disqualification_clauses', []))}")
                    print(f"   技术要求数量: {len(data.get('technical_requirements', []))}")
                    print(f"   资质要求数量: {len(data.get('qualification_requirements', []))}")
                    
                    # 显示部分废标条款
                    if data.get('disqualification_clauses'):
                        print("\n🔍 部分废标条款:")
                        for i, clause in enumerate(data['disqualification_clauses'][:3], 1):
                            print(f"   {i}. {clause.get('content', '未知')[:100]}...")
                            
                    print("\n🎉 测试完成！招标文件分析功能正常运行。")
                else:
                    print("   ⚠️ 未找到分析结果数据")
            else:
                print("   ❌ 分析失败:")
                print(f"   错误信息: {analysis_result.get('error', '未知错误')}")
                
        else:
            print(f"   ❌ 文件上传失败: {result.get('error', '未知错误')}")
            
    except Exception as e:
        print(f"❌ 测试过程中出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_tender_analysis()
