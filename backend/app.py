from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import uuid
import json
from datetime import datetime
from qwen_service import QwenAnalysisService
from file_handler import FileHandler
from database import DatabaseManager

app = Flask(__name__)
CORS(app)  # 允许跨域请求

# 配置
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# 确保上传目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# 初始化服务
qwen_service = QwenAnalysisService()
file_handler = FileHandler()
db_manager = DatabaseManager()

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """
    文件上传接口
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
        
        # 验证文件类型
        if not file_handler.is_allowed_file(file.filename):
            return jsonify({'error': '不支持的文件类型'}), 400
        
        # 生成唯一文件ID
        file_id = str(uuid.uuid4())
        
        # 保存文件
        file_path = file_handler.save_file(file, file_id, app.config['UPLOAD_FOLDER'])
        
        # 提取文件内容
        content = file_handler.extract_content(file_path)
        
        # 保存到数据库
        db_manager.save_file_record(file_id, file.filename, file_path, content)
        
        return jsonify({
            'file_id': file_id,
            'filename': file.filename,
            'message': '文件上传成功'
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/analyze/tender', methods=['POST'])
def analyze_tender():
    """
    招标文件分析接口
    """
    try:
        data = request.get_json()
        file_id = data.get('file_id')
        
        if not file_id:
            return jsonify({'error': '缺少文件ID'}), 400
        
        # 从数据库获取文件内容
        file_record = db_manager.get_file_record(file_id)
        if not file_record:
            return jsonify({'error': '文件不存在'}), 404
        
        # 使用Qwen分析招标文件
        analysis_result = qwen_service.analyze_tender_document(file_record['content'])
        
        # 保存分析结果
        analysis_id = db_manager.save_tender_analysis(file_id, analysis_result)
        
        return jsonify({
            'analysis_id': analysis_id,
            'result': analysis_result,
            'message': '招标文件分析完成'
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/analyze/bid', methods=['POST'])
def analyze_bid():
    """
    投标文件分析接口
    """
    try:
        data = request.get_json()
        file_id = data.get('file_id')
        tender_analysis_id = data.get('tender_analysis_id')
        
        if not file_id:
            return jsonify({'error': '缺少文件ID'}), 400
        
        # 从数据库获取文件内容
        file_record = db_manager.get_file_record(file_id)
        if not file_record:
            return jsonify({'error': '文件不存在'}), 404
        
        # 获取招标文件的分析结果（如果有）
        tender_analysis = None
        if tender_analysis_id:
            tender_analysis = db_manager.get_tender_analysis(tender_analysis_id)
        
        # 使用Qwen分析投标文件
        analysis_result = qwen_service.analyze_bid_document(
            file_record['content'], 
            tender_analysis
        )
        
        # 保存分析结果
        analysis_id = db_manager.save_bid_analysis(file_id, analysis_result)
        
        return jsonify({
            'analysis_id': analysis_id,
            'result': analysis_result,
            'message': '投标文件分析完成'
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/analysis/<analysis_id>', methods=['GET'])
def get_analysis_result(analysis_id):
    """
    获取分析结果接口
    """
    try:
        result = db_manager.get_analysis_result(analysis_id)
        if not result:
            return jsonify({'error': '分析结果不存在'}), 404
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    """
    健康检查接口
    """
    return jsonify({'status': 'healthy', 'timestamp': datetime.now().isoformat()})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
