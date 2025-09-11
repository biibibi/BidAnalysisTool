import requests
import os
import json

def test_text_upload():
    """测试文本文件上传功能"""
    url = 'http://127.0.0.1:5000/api/upload'
    file_path = 'test_upload.txt'
    
    if not os.path.exists(file_path):
        print(f"测试文件不存在: {file_path}")
        return
    
    try:
        with open(file_path, 'rb') as f:
            files = {'file': (file_path, f, 'text/plain')}
            response = requests.post(url, files=files)
            
        print(f"HTTP状态码: {response.status_code}")
        if response.status_code == 200:
            result = response.json()
            print("上传成功!")
            print(f"文件ID: {result.get('file_id')}")
            print(f"文件名: {result.get('filename')}")
            print(f"文件大小: {result.get('size')} 字节")
            print(f"文件类型: {result.get('file_type')}")
            
            # 测试内容提取
            content = result.get('content', '')
            if content:
                print(f"内容长度: {len(content)} 字符")
                print(f"内容预览: {content[:100]}...")
            else:
                print("未能提取到文件内容")
                
        else:
            print(f"上传失败: {response.text}")
            
    except Exception as e:
        print(f"测试过程中出现错误: {e}")

if __name__ == "__main__":
    test_text_upload()
