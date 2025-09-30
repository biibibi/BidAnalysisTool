#!/usr/bin/env python3
"""
投标文件合规性检查工具 - Flask Web应用主文件
=============================================

说明：原文件头在一次补丁时被破坏，现已恢复。该文件提供所有后端 REST 接口。
"""

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os, uuid
from datetime import datetime
from dotenv import load_dotenv

from qwen_service import QwenAnalysisService
from file_handler import FileHandler
from database import DatabaseManager
from ai_agents.agent_manager import agent_manager
from ai_agents.document_processor import document_processor

load_dotenv()

app = Flask(__name__, static_folder="../static", static_url_path="/static")
CORS(app)

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
uploads_path = os.path.join(project_root, 'uploads')
os.makedirs(uploads_path, exist_ok=True)

app.config['UPLOAD_FOLDER'] = uploads_path
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

qwen_service = QwenAnalysisService()
file_handler = FileHandler()
db_manager = DatabaseManager()

# ---------- 工具函数 ----------
def handle_api_error(e, default_message="操作失败"):
    return jsonify({'error': str(e) if str(e) else default_message}), 500

def validate_file_id(file_id, error_message="缺少文件ID"):
    if not file_id:
        return jsonify({'error': error_message}), 400
    return None

def get_file_record_or_error(file_id):
    rec = db_manager.get_file_record(file_id)
    if not rec:
        return None, (jsonify({'error': '文件不存在'}), 404)
    return rec, None

@app.route('/')
def index():
    return send_from_directory('../frontend', 'index.html')

@app.route('/favicon.ico')
def favicon():
    return app.send_static_file('icon/Unicom1.png')

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """文件上传接口（恢复版）"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '未找到上传文件字段 file'}), 400
        f = request.files['file']
        if f.filename == '':
            return jsonify({'error': '文件名为空'}), 400
        file_id = str(uuid.uuid4())
        filename = f.filename
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], file_id + '_' + filename)
        f.save(save_path)
        try:
            content = file_handler.extract_content(save_path)
        except Exception:
            content = ''
        db_manager.save_file_record(file_id, filename, save_path, content)
        return jsonify({'file_id': file_id, 'filename': filename, 'message': '文件上传成功'})
    except Exception as e:
        return handle_api_error(e)
        file_size = os.path.getsize(temp_file_path)
        
        # 检查是否存在重复文件（仅基于文件大小判断）
        duplicate_file = db_manager.find_duplicate_file(file_size)
        if duplicate_file:
            # 删除刚上传的临时文件，因为已存在相同大小的文件
            os.remove(temp_file_path)
            
            # 返回已存在文件的信息
            return jsonify({
                'file_id': duplicate_file['id'],
                'filename': duplicate_file['filename'],
                'message': f'检测到相同大小的文件已存在，直接使用现有文件（大小: {file_size} 字节）',
                'is_duplicate': True,
                'original_upload_time': duplicate_file['upload_time'],
                'file_size': file_size
            })
        
        # 文件不重复，继续处理
        file_path = temp_file_path
        
        # 提取文件内容（文本内容）
        content = file_handler.extract_content(file_path)
        
        # 将文件记录保存到数据库
        db_manager.save_file_record(file_id, filename, file_path, content)
        
        # 返回成功响应
        return jsonify({
            'file_id': file_id,
            'filename': filename,
            'message': '文件上传成功'
        })
        
    except Exception as e:
        # 捕获并返回所有异常
        return handle_api_error(e)

@app.route('/api/analyze/tender', methods=['POST'])
def analyze_tender():
    """
    招标文件分析接口
    ================
    
    对上传的招标文件进行AI智能分析，提取废标条款和关键要求。
    该接口会调用Qwen大模型来识别可能导致投标无效的条款。
    
    请求方式：POST
    请求头：Content-Type: application/json
    请求参数：
        {
            "file_id": "已上传文件的ID"
        }
    
    响应格式：
        成功: {
            "analysis_id": "分析结果ID",
            "result": {
                "summary": "分析摘要",
                "invalid_items": [
                    {
                        "category": "条款类别",
                        "description": "具体描述",
                        "requirement": "具体要求",
                        "severity": "严重程度",
                        "keywords": ["关键词"]
                    }
                ],
                "suggestions": ["建议列表"]
            },
            "message": "招标文件分析完成"
        }
        失败: {
            "error": "错误信息"
        }
    
    分析内容包括：
        - 资质要求
        - 技术要求  
        - 商务要求
        - 格式要求
        - 时间要求
        - 其他废标条款
    
    Returns:
        JSON响应，包含分析结果或错误信息
        
    Raises:
        400: 请求参数错误（缺少文件ID）
        404: 文件不存在
        500: 服务器内部错误或AI分析失败
    """
    try:
        # 解析JSON请求数据
        data = request.get_json()
        file_id = data.get('file_id')
        # 默认从环境变量读取全局提供方（可选：'qwen' 或 'doubao'）
        provider = os.getenv('LLM_PROVIDER', 'qwen')
        
        # 验证必需参数
        error_response = validate_file_id(file_id)
        if error_response:
            return error_response
        
        # 从数据库获取文件记录和内容
        file_record, error_response = get_file_record_or_error(file_id)
        if error_response:
            return error_response
        if not file_record:
            return jsonify({'error': '文件不存在'}), 404
        
        # 使用AI分析招标文件内容（可选模型路由，默认Qwen）
        if provider:
            analysis_result = qwen_service.analyze_tender_document_with_model(
                file_record['content'], provider=provider
            )
        else:
            analysis_result = qwen_service.analyze_tender_document(file_record['content'])
        
        # 将分析结果保存到数据库
        analysis_id = db_manager.save_tender_analysis(file_id, analysis_result)
        
        # 返回分析结果
        return jsonify({
            'analysis_id': analysis_id,
            'result': analysis_result,
            'message': '招标文件分析完成'
        })
        
    except Exception as e:
        # 捕获并返回所有异常
        return handle_api_error(e)

@app.route('/api/analyze/bid', methods=['POST'])
def analyze_bid():
    """
    投标文件分析接口
    ================
    
    对上传的投标文件进行合规性检查，识别潜在的废标风险。
    该接口会基于招标文件的要求（如果提供）来检查投标文件的合规性。
    
    请求方式：POST
    请求头：Content-Type: application/json
    请求参数：
        {
            "file_id": "已上传投标文件的ID",
            "tender_analysis_id": "招标文件分析结果ID（可选）"
        }
    
    响应格式：
        成功: {
            "analysis_id": "分析结果ID",
            "result": {
                "summary": "合规性分析摘要",
                "compliance_check": {
                    "overall_status": "合规状态",
                    "risk_level": "风险等级",
                    "score": "合规性得分"
                },
                "issues": [
                    {
                        "category": "问题类别",
                        "description": "问题描述",
                        "severity": "严重程度",
                        "suggestion": "改进建议",
                        "location": "问题位置"
                    }
                ],
                "recommendations": ["改进建议列表"]
            },
            "message": "投标文件分析完成"
        }
        失败: {
            "error": "错误信息"
        }
    
    检查内容包括：
        - 资质要求符合性
        - 技术规格匹配度
        - 商务条件满足情况
        - 文件格式正确性
        - 必要信息完整性
        - 签名盖章完整性
    
    Returns:
        JSON响应，包含合规性检查结果或错误信息
        
    Raises:
        400: 请求参数错误（缺少文件ID）
        404: 文件不存在
        500: 服务器内部错误或AI分析失败
    """
    try:
        # 解析JSON请求数据
        data = request.get_json()
        file_id = data.get('file_id')
        tender_analysis_id = data.get('tender_analysis_id')
        # 请求可显式指定；未指定则使用全局默认提供方
        provider = data.get('provider') or os.getenv('LLM_PROVIDER', 'qwen')  # 可选：'qwen' 或 'doubao'
        
        # 验证必需参数
        error_response = validate_file_id(file_id)
        if error_response:
            return error_response
        
        # 从数据库获取投标文件记录和内容
        file_record, error_response = get_file_record_or_error(file_id)
        if error_response:
            return error_response
        if not file_record:
            return jsonify({'error': '文件不存在'}), 404
        
        # 获取招标文件的分析结果（如果提供了分析ID）
        # 这将用于更精确的合规性检查
        tender_analysis = None
        if tender_analysis_id:
            tender_analysis = db_manager.get_tender_analysis(tender_analysis_id)
        
        # 进行投标文件合规性分析（可选模型路由，默认Qwen）
        if provider:
            analysis_result = qwen_service.analyze_bid_document_with_model(
                file_record['content'], tender_analysis, provider=provider
            )
        else:
            analysis_result = qwen_service.analyze_bid_document(
                file_record['content'], tender_analysis
            )
        
        # 将分析结果保存到数据库
        analysis_id = db_manager.save_bid_analysis(file_id, analysis_result, tender_analysis_id)
        
        # 返回分析结果
        return jsonify({
            'analysis_id': analysis_id,
            'result': analysis_result,
            'message': '投标文件分析完成'
        })
        
    except Exception as e:
        # 捕获并返回所有异常
        return handle_api_error(e)

@app.route('/api/process-bid-document', methods=['POST'])
def process_bid_document():
    """
    投标文件预处理接口
    ==================
    
    对投标文件进行预处理，包括目录提取和文档拆分。
    在正式分析之前调用，生成目录结构和拆分文档。
    
    请求方式：POST
    请求头：Content-Type: application/json
    请求参数：
        {
            "file_id": "已上传投标文件的ID"
        }
    
    响应格式：
        成功: {
            "success": true,
            "work_dir": "工作目录路径",
            "toc_result": {目录提取结果},
            "split_result": {文档拆分结果},
            "message": "文档处理完成"
        }
        失败: {
            "success": false,
            "error": "错误信息"
        }
    
    Returns:
        JSON响应，包含处理结果或错误信息
    """
    try:
        # 解析JSON请求数据
        data = request.get_json()
        file_id = data.get('file_id')
        
        # 验证必需参数
        error_response = validate_file_id(file_id)
        if error_response:
            return error_response
        
        # 从数据库获取投标文件记录和路径
        file_record, error_response = get_file_record_or_error(file_id)
        if error_response:
            return error_response
        if not file_record:
            return jsonify({'success': False, 'error': '文件不存在'}), 404
        
        file_path = file_record['file_path']
        if not os.path.exists(file_path):
            return jsonify({'success': False, 'error': '文件路径不存在'}), 404
        
        # 检查文件格式
        if not file_path.lower().endswith(('.docx', '.doc')):
            return jsonify({'success': False, 'error': '仅支持Word文档格式(.docx/.doc)'}), 400
        
        # 调用文档处理器
        result = document_processor.process_bid_document(file_path, file_id)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 500
            
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'文档处理异常: {str(e)}'
        }), 500

@app.route('/api/process-bid-document/status/<file_id>', methods=['GET'])
def get_bid_document_status(file_id):
    """
    获取投标文件处理状态接口
    ========================
    
    查询指定文件ID的文档处理状态。
    
    请求方式：GET
    URL参数：
        file_id: 文件ID
    
    响应格式：
        {
            "exists": true/false,
            "work_dir": "工作目录路径",
            "has_toc": true/false,
            "toc_files": ["目录文件列表"],
            "has_splits": true/false,
            "split_files": ["拆分文件列表"],
            "split_count": 拆分文件数量
        }
    
    Returns:
        JSON响应，包含处理状态信息
    """
    try:
        # 验证文件ID
        error_response = validate_file_id(file_id)
        if error_response:
            return error_response
        
        # 获取处理状态
        status = document_processor.get_processing_status(file_id)
        
        return jsonify(status)
        
    except Exception as e:
        return handle_api_error(e)

@app.route('/api/analysis/<analysis_id>', methods=['GET'])
def get_analysis_result(analysis_id):
    """
    获取分析结果接口
    ================
    
    根据分析ID获取之前保存的分析结果。
    支持获取招标文件分析结果和投标文件分析结果。
    
    请求方式：GET
    URL参数：
        analysis_id: 分析结果的唯一标识符
    
    响应格式：
        成功: {
            "id": "分析ID",
            "type": "分析类型(tender/bid)",
            "file_id": "关联的文件ID", 
            "analysis_result": {分析结果对象},
            "created_time": "创建时间"
        }
        失败: {
            "error": "错误信息"
        }
    
    使用场景：
        - 查看历史分析结果
        - 前端页面刷新后恢复数据
        - 分享分析结果链接
        - 导出分析报告
    
    Returns:
        JSON响应，包含完整的分析结果或错误信息
        
    Raises:
        404: 分析结果不存在
        500: 服务器内部错误
    """
    try:
        # 从数据库查询分析结果
        # 自动判断是招标文件分析还是投标文件分析
        result = db_manager.get_analysis_result(analysis_id)
        
        # 检查结果是否存在
        if not result:
            return jsonify({'error': '分析结果不存在'}), 404
        
        # 返回完整的分析结果
        return jsonify(result)
        
    except Exception as e:
        # 捕获并返回所有异常
        return handle_api_error(e)


@app.route('/api/health', methods=['GET'])
def health_check():
    """
    健康检查接口
    ============
    
    用于检查服务器和API的运行状态。
    主要用于监控、负载均衡器健康检查和服务状态确认。
    
    请求方式：GET
    无需参数
    
    响应格式：
        {
            "status": "healthy",
            "timestamp": "当前时间戳(ISO格式)"
        }
    
    使用场景：
        - 系统监控
        - 负载均衡器健康检查
        - 部署后状态确认
        - API可用性测试
    
    Returns:
        JSON响应，包含健康状态和时间戳
    """
    return jsonify({
        'status': 'healthy', 
        'timestamp': datetime.now().isoformat()
    })


@app.route('/api/check-project-info', methods=['POST'])
def check_project_info():
    """
    项目信息检测接口
    ================
    
    专门用于检测投标文件中的项目名称、项目编号等信息是否与招标文件一致。
    这是针对"项目名称、项目编号等内容是否正确，内容是否和招标文件要求一致"检测项的专用接口。
    
    请求方式：POST
    请求头：Content-Type: application/json
    请求参数：
        {
            "bid_file_id": "投标文件ID",
            "tender_file_id": "招标文件ID（可选）",
            "check_type": "检测类型（固定为project_info）"
        }
    
    响应格式：
        成功: {
            "success": true,
            "data": {
                "has_errors": true/false,
                "error_count": 错误数量,
                "errors": [
                    {
                        "type": "错误类型",
                        "found_value": "发现的值",
                        "correct_value": "正确的值",
                        "location": "位置信息",
                        "context": "上下文",
                        "severity": "严重程度",
                        "description": "错误描述"
                    }
                ],
                "confidence": 0.95,
                "tender_info": {
                    "project_id": "招标项目编号",
                    "project_name": "招标项目名称"
                },
                "bid_info": {
                    "project_id": "投标项目编号",
                    "project_name": "投标项目名称"
                }
            },
            "message": "项目信息检测完成"
        }
        失败: {
            "success": false,
            "error": "错误信息"
        }
    
    Returns:
        JSON响应，包含项目信息检测结果或错误信息
    """
    try:
        data = request.get_json()
        bid_file_id = data.get('bid_file_id')
        tender_file_id = data.get('tender_file_id')
        check_type = data.get('check_type')
        
        # 验证必需参数
        if not bid_file_id:
            return jsonify({'success': False, 'error': '缺少投标文件ID'}), 400
        
        if check_type != 'project_info':
            return jsonify({'success': False, 'error': '不支持的检测类型'}), 400
        
        # 获取投标文件
        bid_file, error_response = get_file_record_or_error(bid_file_id)
        if error_response:
            return jsonify({'success': False, 'error': '投标文件不存在'}), 404
        if not bid_file:
            return jsonify({'success': False, 'error': '投标文件不存在'}), 404
        if not bid_file:
            return jsonify({'success': False, 'error': '投标文件不存在'}), 404
        
        # 获取招标文件信息（如果提供了招标文件ID）
        tender_info = None
        if tender_file_id:
            tender_file, tender_error = get_file_record_or_error(tender_file_id)
            if tender_error:
                return jsonify({'success': False, 'error': '招标文件不存在'}), 404
            if not tender_file:
                return jsonify({'success': False, 'error': '招标文件不存在'}), 404
            if not tender_file:
                return jsonify({'success': False, 'error': '招标文件不存在'}), 404
            
            # 从招标文件提取项目信息
            tender_extract_result = agent_manager.extract_project_info(
                tender_file['content'], 
                'tender',
                file_id=tender_file_id
            )
            
            if tender_extract_result.get('success'):
                tender_info = {
                    'project_id': tender_extract_result['data'].get('project_id'),
                    'project_name': tender_extract_result['data'].get('project_name')
                }
            else:
                return jsonify({
                    'success': False, 
                    'error': '招标文件项目信息提取失败: ' + tender_extract_result.get('error', '未知错误')
                }), 400
        
        # 如果没有招标文件信息，返回错误
        if not tender_info or (not tender_info.get('project_id') and not tender_info.get('project_name')):
            return jsonify({
                'success': False, 
                'error': '缺少招标文件项目信息，无法进行对比检测'
            }), 400
        
        # 使用ProjectInfoAgent进行错误检测
        context = {
            'document_type': 'bid',
            'tender_project_id': tender_info.get('project_id'),
            'tender_project_name': tender_info.get('project_name')
        }
        
        # 调用agent进行检测
        result = agent_manager.process_with_agent(
            'ProjectInfoAgent', 
            bid_file['content'], 
            context
        )
        
        if not result.get('success'):
            return jsonify({
                'success': False,
                'error': '项目信息检测失败: ' + result.get('error', '未知错误')
            }), 500
        
        # 处理检测结果
        detection_data = result['data']
        
        # 获取从错误检测过程中同时提取的投标文件信息（避免重复AI调用）
        bid_info = detection_data.get('bid_info', {})
        
        # 整合检测结果
        response_data = {
            'has_errors': detection_data.get('has_errors', False),
            'error_count': detection_data.get('error_count', 0),
            'errors': detection_data.get('errors', []),
            'confidence': detection_data.get('confidence', 0.8),  # 使用Agent内部计算的置信度
            'tender_info': tender_info,
            'bid_info': bid_info,
            'detection_details': detection_data
        }
        
        return jsonify({
            'success': True,
            'data': response_data,
            'message': '项目信息检测完成'
        })
        
    except Exception as e:
        return handle_api_error(e)


@app.route('/api/extract-project-info', methods=['POST'])
def extract_project_info():
    """
    项目信息提取接口
    ================
    
    从招标文件或投标文件中提取项目编号和项目名称。
    支持智能识别文档类型并提取相应的项目信息。
    
    请求方式：POST
    请求头：Content-Type: application/json
    请求参数：
        {
            "file_id": "已上传文件的ID",
            "document_type": "文档类型(tender/bid/auto)"
        }
    
    响应格式：
        成功: {
            "success": true,
            "data": {
                "project_id": "项目编号",
                "project_name": "项目名称",
                "document_type": "文档类型",
                "confidence": "提取置信度",
                "extraction_methods": {提取方法详情}
            },
            "message": "提取完成"
        }
        失败: {
            "success": false,
            "error": "错误信息"
        }
    
    Returns:
        JSON响应，包含项目信息或错误信息
    """
    try:
        data = request.get_json()
        file_id = data.get('file_id')
        document_type = data.get('document_type', 'auto')
        
        # 验证必需参数
        error_response = validate_file_id(file_id)
        if error_response:
            return error_response
        
        # 从数据库获取文件内容
        file_record, error_response = get_file_record_or_error(file_id)
        if error_response:
            return error_response
        if not file_record:
            return jsonify({'error': '文件不存在'}), 404
        
        # 使用Agent提取项目信息
        result = agent_manager.extract_project_info(
            file_record['content'], 
            document_type,
            file_id=file_id
        )
        
        # 如果提取成功，返回结果
        return jsonify(result)
        
    except Exception as e:
        return handle_api_error(e)


@app.route('/api/match-project-info', methods=['POST'])
def match_project_info():
    """
    项目信息匹配接口
    ================
    
    将投标文件的项目信息与招标文件的项目信息进行匹配验证。
    检查项目编号和项目名称是否一致。
    
    请求方式：POST
    请求头：Content-Type: application/json
    请求参数：
        {
            "bid_file_id": "投标文件ID",
            "tender_info": {
                "project_id": "招标项目编号",
                "project_name": "招标项目名称"
            }
        }
        或
        {
            "bid_file_id": "投标文件ID",
            "tender_file_id": "招标文件ID"
        }
    
    响应格式：
        {
            "success": true,
            "data": {
                "overall_match": true/false,
                "project_id_match": {
                    "match": true/false,
                    "bid_id": "投标项目编号",
                    "tender_id": "招标项目编号"
                },
                "project_name_match": {
                    "match": true/false,
                    "similarity": 0.95,
                    "bid_name": "投标项目名称",
                    "tender_name": "招标项目名称"
                },
                "analysis": "匹配分析报告"
            }
        }
    
    Returns:
        JSON响应，包含匹配结果或错误信息
    """
    try:
        data = request.get_json()
        bid_file_id = data.get('bid_file_id')
        
        # 验证必需参数
        error_response = validate_file_id(bid_file_id, '缺少投标文件ID')
        if error_response:
            return error_response
        
        # 获取投标文件内容
        bid_file, error_response = get_file_record_or_error(bid_file_id)
        if error_response:
            return error_response
        if not bid_file:
            return jsonify({'error': '投标文件不存在'}), 404
        
        # 获取招标文件信息
        tender_info = data.get('tender_info')
        if not tender_info:
            # 如果没有直接提供，尝试从招标文件ID获取
            tender_file_id = data.get('tender_file_id')
            if tender_file_id:
                tender_file, tender_error = get_file_record_or_error(tender_file_id)
                if tender_error:
                    return tender_error
                if not tender_file:
                    return jsonify({'error': '招标文件不存在'}), 404
                
                # 提取招标文件的项目信息
                tender_result = agent_manager.extract_project_info(
                    tender_file['content'], 
                    'tender',
                    file_id=tender_file_id
                )
                if tender_result.get('success'):
                    tender_info = tender_result['data']
                else:
                    return jsonify({'error': '招标文件项目信息提取失败'}), 400
            else:
                return jsonify({'error': '缺少招标文件信息'}), 400
        
        # 进行项目信息匹配
        match_result = agent_manager.match_project_info(
            bid_file['content'], 
            tender_info
        )
        
        # 返回匹配结果
        return jsonify(match_result)
        
    except Exception as e:
        return handle_api_error(e)


@app.route('/api/analyze/authorization-letter', methods=['POST'])
def analyze_authorization_letter():
    """授权委托书/身份证明核验接口
    =================================

    使用 AuthorizationLetterAgent 对上传的授权委托书（或身份证明）Word 文档进行多模态核验：
        - 法定代表人/被授权人身份信息识别
        - 证件有效性/有效期识别（若存在）
        - 授权信息完整性
        - 项目编号、项目名称与招标文件登记信息一致性

    请求 JSON:
        {
          "file_id": "授权委托书文件ID (必填)",
          "tender_file_id": "招标文件 file_id (可选)",
          "tender_analysis_id": "招标分析记录 ID (可选)",
          "expected_project_id": "外部传入期望项目编号 (可选)",
          "expected_project_name": "外部传入期望项目名称 (可选)",
          "provider": "qwen|doubao (可选, 默认环境变量 LLM_PROVIDER)"
        }

    响应示例(success):
        {
          "success": true,
          "data": { ... AuthorizationLetterAgent 返回的数据 ... },
          "message": "授权委托书多模态核验完成"
        }

    错误:
        400 缺少参数 / 文件不存在
        500 内部错误
    """
    try:
        data = request.get_json() or {}
        file_id = data.get('file_id')
        if not file_id:
            return jsonify({'success': False, 'error': '缺少 file_id'}), 400

        file_record, error_response = get_file_record_or_error(file_id)
        if error_response:
            # error_response 已包含 (json, status_code)
            return error_response
        if not file_record:
            return jsonify({'success': False, 'error': '文件不存在'}), 404

        file_path = file_record.get('file_path')
        if not file_path or not os.path.exists(file_path):
            return jsonify({'success': False, 'error': '文件路径不存在'}), 404
        if not file_path.lower().endswith(('.docx', '.doc')):
            return jsonify({'success': False, 'error': '仅支持 Word 文档(.doc/.docx)'}), 400

        # 自动选择授权委托书拆分文档（可选）
        selected_file_path = file_path
        selection_meta = None
        if data.get('auto_select'):
            try:
                # 推断work_dir: temp/<file_id>/split_documents
                current_dir = os.path.dirname(os.path.abspath(__file__))
                project_root = os.path.dirname(current_dir)
                temp_dir = os.path.join(project_root, 'temp', file_id)
                split_dir = os.path.join(temp_dir, 'split_documents')
                toc_md = None
                # 查找目录md
                if os.path.isdir(temp_dir):
                    for item in os.listdir(temp_dir):
                        if item.endswith('_目录.md'):
                            toc_md = os.path.join(temp_dir, item)
                            break
                candidates = []
                if os.path.isdir(split_dir):
                    for fname in os.listdir(split_dir):
                        if fname.lower().endswith(('.docx', '.doc')):
                            fpath = os.path.join(split_dir, fname)
                            norm = fname.lower()
                            score = 0
                            # 关键字打分
                            KEYWORDS = [
                                '授权', '委托', '委托书', '身份证明', '法定代表', '身份证', '代表人', '授权书'
                            ]
                            for kw in KEYWORDS:
                                if kw in norm:
                                    score += 3
                            # 文件名长度过长扣分
                            if len(fname) > 80:
                                score -= 1
                            candidates.append({'file': fpath, 'name': fname, 'score': score})
                chosen = None
                if candidates:
                    # 读取 toc md 进一步加权
                    toc_text = ''
                    if toc_md and os.path.exists(toc_md):
                        try:
                            with open(toc_md, 'r', encoding='utf-8', errors='ignore') as f:
                                toc_text = f.read().lower()
                        except Exception:
                            toc_text = ''
                    if toc_text:
                        for c in candidates:
                            base = os.path.splitext(c['name'])[0].lower()
                            if base and base in toc_text:
                                c['score'] += 2
                            # 章节行若含“授权委托书”并且序号靠前权重更大
                            if '授权' in toc_text and '委托' in toc_text:
                                c['score'] += 1
                    candidates.sort(key=lambda x: x['score'], reverse=True)
                    # 过滤掉完全没有匹配分数的
                    filtered = [c for c in candidates if c['score'] > 0]
                    chosen = filtered[0] if filtered else candidates[0]
                if chosen:
                    selected_file_path = chosen['file']
                    selection_meta = {
                        'strategy': 'auto_select',
                        'chosen_name': os.path.basename(chosen['file']),
                        'score': chosen['score'],
                        'candidate_count': len(candidates)
                    }
                else:
                    selection_meta = {
                        'strategy': 'fallback_original',
                        'reason': '无可用拆分候选, 使用原始文件'
                    }
            except Exception as sel_exc:
                selection_meta = {
                    'strategy': 'fallback_error',
                    'error': str(sel_exc)
                }

        # 组装上下文
        context = {
            'tender_file_id': data.get('tender_file_id'),
            'tender_analysis_id': data.get('tender_analysis_id'),
            'expected_project_id': data.get('expected_project_id'),
            'expected_project_name': data.get('expected_project_name'),
            'provider': data.get('provider') or os.getenv('LLM_PROVIDER', 'qwen')
        }

        result = agent_manager.process_with_agent('AuthorizationLetterAgent', selected_file_path, context)
        if selection_meta:
            result.setdefault('data', {})
            if isinstance(result.get('data'), dict):
                result['data']['selected_file'] = selected_file_path
                result['data']['selection_meta'] = selection_meta

        status_code = 200 if result.get('success') else 500
        return jsonify(result), status_code
    except Exception as e:
        return handle_api_error(e, '授权委托书核验失败')


@app.route('/api/agents', methods=['GET'])
def list_agents():
    """
    列出所有可用的AI Agent
    ======================
    
    返回系统中所有已注册的AI Agent信息。
    
    请求方式：GET
    无需参数
    
    响应格式：
        {
            "agents": [
                {
                    "name": "Agent名称",
                    "description": "Agent描述",
                    "type": "Agent类型"
                }
            ],
            "statistics": {
                "total_agents": 数量,
                "agent_names": ["名称列表"]
            }
        }
    
    Returns:
        JSON响应，包含Agent列表和统计信息
    """
    try:
        agents = agent_manager.list_agents()
        stats = agent_manager.get_agent_statistics()
        
        return jsonify({
            'agents': agents,
            'statistics': stats
        })
        
    except Exception as e:
        return handle_api_error(e)


@app.route('/api/preview/<file_id>', methods=['GET'])
def preview_file(file_id):
    """
    文件预览接口
    ============
    
    提供文件内容的在线预览功能。根据文件类型返回适当的预览格式。
    
    请求方式：GET
    路径参数：
        file_id: 文件ID
    
    响应格式：
        成功: 返回文件内容（HTML格式用于预览）
        失败: JSON错误信息
    """
    try:
        # 验证文件ID
        error_response = validate_file_id(file_id)
        if error_response:
            return error_response
        
        # 获取文件记录
        file_record, error_response = get_file_record_or_error(file_id)
        if error_response:
            return error_response
        if not file_record:
            return jsonify({'error': '文件不存在'}), 404
        if not file_record:
            return jsonify({'error': '文件不存在'}), 404
        
        # 生成预览HTML
        html_content = f"""
        <!DOCTYPE html>
        <html lang="zh-CN">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>文件预览 - {file_record['filename']}</title>
            <style>
                body {{
                    font-family: 'Microsoft YaHei', Arial, sans-serif;
                    line-height: 1.6;
                    margin: 0;
                    padding: 20px;
                    background-color: #f5f5f5;
                }}
                .container {{
                    max-width: 800px;
                    margin: 0 auto;
                    background: white;
                    padding: 30px;
                    border-radius: 10px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                }}
                .header {{
                    border-bottom: 2px solid #007bff;
                    padding-bottom: 15px;
                    margin-bottom: 25px;
                }}
                .filename {{
                    color: #007bff;
                    font-size: 1.5em;
                    font-weight: bold;
                    margin: 0;
                }}
                .meta {{
                    color: #666;
                    font-size: 0.9em;
                    margin-top: 5px;
                }}
                .content {{
                    background: #f8f9fa;
                    padding: 20px;
                    border-radius: 5px;
                    border-left: 4px solid #007bff;
                    white-space: pre-wrap;
                    word-wrap: break-word;
                    max-height: 600px;
                    overflow-y: auto;
                }}
                .footer {{
                    margin-top: 25px;
                    padding-top: 15px;
                    border-top: 1px solid #dee2e6;
                    color: #666;
                    font-size: 0.85em;
                    text-align: center;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1 class="filename">{file_record['filename']}</h1>
                    <div class="meta">
                        上传时间: {file_record['upload_time']} | 
                        文件大小: {file_record['file_size']} 字节
                    </div>
                </div>
                <div class="content">{file_record['content']}</div>
                <div class="footer">
                    <p>BidAnalysis Tool - 文件预览服务</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return html_content, 200, {'Content-Type': 'text/html; charset=utf-8'}
        
    except Exception as e:
        return handle_api_error(e)


@app.route('/api/download/<file_id>', methods=['GET'])
def download_file(file_id):
    """
    文件下载接口
    ============
    
    提供文件下载功能，返回原始文件内容。
    
    请求方式：GET
    路径参数：
        file_id: 文件ID
    
    响应格式：
        成功: 文件内容（attachment形式下载）
        失败: JSON错误信息
    """
    try:
        # 验证文件ID
        error_response = validate_file_id(file_id)
        if error_response:
            return error_response
        
        # 获取文件记录
        file_record, error_response = get_file_record_or_error(file_id)
        if error_response:
            return error_response
        if not file_record:
            return jsonify({'error': '文件不存在'}), 404
        
        # 确定文件类型
        filename = file_record['filename']
        if filename.lower().endswith('.pdf'):
            mimetype = 'application/pdf'
        elif filename.lower().endswith(('.doc', '.docx')):
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        elif filename.lower().endswith('.txt'):
            mimetype = 'text/plain'
        else:
            mimetype = 'application/octet-stream'
        
        # 创建响应
        response = app.response_class(
            file_record['content'],
            mimetype=mimetype,
            headers={
                'Content-Disposition': f'attachment; filename="{filename}"',
                'Content-Length': str(len(file_record['content']))
            }
        )
        
        return response
        
    except Exception as e:
        return handle_api_error(e)


# === 应用启动配置 ===
if __name__ == '__main__':
    """
    应用启动入口
    ============
    
    当直接运行此文件时启动Flask开发服务器。
    生产环境建议使用专业的WSGI服务器如Gunicorn。
    
    配置参数：
        debug=True: 启用调试模式，代码变更时自动重载
        host='0.0.0.0': 监听所有网络接口，允许外部访问
        port=5000: 监听端口号
    
    注意：
        调试模式不应在生产环境中使用，存在安全风险。
    """
    app.run(debug=True, host='0.0.0.0', port=5000)
