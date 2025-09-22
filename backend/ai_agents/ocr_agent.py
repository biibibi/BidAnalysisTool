#!/usr/bin/env python3
"""
OCR智能识别代理
===============

本模块实现基于大模型的OCR图像智能识别功能。
支持多种AI模型（Qwen、Doubao等），提供证书日期验证和身份证信息提取等核心功能。

核心功能：
    1. 证书/授权函/社保证明日期有效性验证
    2. 身份证信息提取（姓名、性别、身份证号）
    3. 通用OCR文本识别

技术特点：
    - 继承BaseAgent，遵循统一的Agent接口
    - 支持多种AI模型提供方（Qwen、Doubao等）
    - 自动图像预处理和格式转换
    - 智能日期解析和有效性判断
    - 结构化的识别结果输出

依赖库：
    - openai: OpenAI Python SDK
    - PIL: Python图像处理库
    - base64: 图像编码处理
    - re: 正则表达式
    - datetime: 日期时间处理

环境要求：
    - DASHSCOPE_API_KEY: 阿里云百炼API密钥
    - ARK_API_KEY: 字节方舟API密钥（可选）

作者：BidAnalysis Team
创建时间：2025年
版本：1.0
"""

import os
import json
import base64
import re
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List, cast
from PIL import Image
import io

from openai import OpenAI
from dotenv import load_dotenv
from .base_agent import BaseAgent

# 确保在独立脚本/测试环境下也能读取 backend/.env
_env_loaded = False
try:
    _backend_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _env_path = os.path.join(_backend_dir, '.env')
    if os.path.exists(_env_path):
        load_dotenv(_env_path, override=False)
        _env_loaded = True
except Exception:
    # 静默失败，不影响后续从系统环境读取
    _env_loaded = False


class OCRAgent(BaseAgent):
    """
    OCR智能识别代理类
    ==================
    
    基于大模型的图像OCR识别和智能分析代理。
    支持证书日期验证、身份证信息提取等专业功能。
    
    主要方法：
        - validate_document_date(): 验证证书/授权函等文档日期有效性
        - extract_id_card_info(): 提取身份证信息
        - _call_vision_api(): 调用视觉AI API
        - _preprocess_image(): 图像预处理
        
    使用示例：
        agent = OCRAgent()
        result = agent.validate_document_date(image_path)
    """
    
    def __init__(self, default_provider: str = "qwen"):
        """
        初始化OCR代理
        
        Args:
            default_provider (str): 默认使用的AI模型提供方
                支持值：'qwen', 'doubao', 'ali', 'dashscope', 'ark', 'volc'
        """
        super().__init__(
            name="OCR智能识别代理",
            description="基于大模型的图像OCR识别和智能分析代理，支持证书日期验证和身份证信息提取"
        )
        
        # 设置默认模型提供方
        self.default_provider = (default_provider or os.getenv("LLM_PROVIDER", "qwen")).lower()
        
        # 初始化Qwen客户端
        self.qwen_client = OpenAI(
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
        )
        self.qwen_model = "qwen-vl-plus"  # 支持视觉的Qwen模型
        
        # 初始化豆包（方舟）客户端配置
        self.ark_base_url = os.getenv("ARK_BASE_URL", "https://ark.cn-beijing.volces.com/api/v3")
        self.doubao_model_id = os.getenv("DOUBAO_MODEL_ID", "doubao-vision-32k")  # 支持视觉的豆包模型
        
        self.logger.info(f"OCR Agent初始化完成，默认使用模型提供方: {self.default_provider}")
    
    def process(self, content: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        处理OCR识别请求
        
        Args:
            content (str): 图像文件路径或base64编码的图像数据
            context (Optional[Dict[str, Any]]): 可选的上下文信息
                - task_type (str): 任务类型 ('date_validation', 'id_extraction', 'general_ocr')
                - provider (str): 指定使用的模型提供方
                
        Returns:
            Dict[str, Any]: 处理结果
        """
        if not self.validate_input(content):
            return self.create_error_result("输入内容无效")
        
        try:
            # 获取任务类型和模型提供方
            task_type = context.get("task_type", "general_ocr") if context else "general_ocr"
            provider = context.get("provider", self.default_provider) if context else self.default_provider
            
            self.log_processing(len(content), context)
            
            # 根据任务类型调用相应的处理方法
            if task_type == "date_validation":
                return self.validate_document_date(content, provider)
            elif task_type == "id_extraction":
                return self.extract_id_card_info(content, provider)
            else:
                return self.general_ocr(content, provider)
                
        except Exception as e:
            self.logger.error(f"OCR处理失败: {str(e)}")
            return self.create_error_result(f"OCR处理失败: {str(e)}")
    
    def validate_document_date(self, image_input: str, provider: Optional[str] = None) -> Dict[str, Any]:
        """
        验证证书/授权函/社保证明等文档中的日期有效性
        ============================================
        
        使用AI视觉模型识别图像中的日期信息，并判断是否在有效期内。
        
        Args:
            image_input (str): 图像文件路径或base64编码的图像数据
            provider (str): 指定使用的模型提供方
            
        Returns:
            Dict[str, Any]: 验证结果
                - success (bool): 处理是否成功
                - data (dict): 验证结果数据
                    - is_valid (bool): 日期是否有效
                    - dates_found (List[Dict]): 识别到的日期信息
                        - date (str): 日期字符串
                        - type (str): 日期类型（发证日期、有效期等）
                        - is_expired (bool): 是否已过期
                    - document_type (str): 文档类型
                    - confidence (float): 识别置信度
                - message (str): 处理消息
                
        识别目标：
            - 营业执照有效期
            - 资质证书有效期
            - 授权委托书日期
            - 社保证明时间范围
            - 其他证明文件日期
        """
        try:
            provider = provider or self.default_provider
            
            # 构建日期验证提示词
            prompt = """
请分析这张图片，识别其中的日期信息并判断有效性。

请重点关注以下类型的日期：
1. 证书有效期（如营业执照、资质证书）
2. 授权委托书的授权日期和有效期
3. 社保证明的时间范围
4. 其他重要的时间信息

请按照以下JSON格式返回分析结果：
{
    "document_type": "文档类型（如：营业执照、授权委托书、社保证明等）",
    "dates_found": [
        {
            "date": "识别到的日期（YYYY-MM-DD格式）",
            "type": "日期类型（如：发证日期、有效期至、授权日期等）",
            "original_text": "图片中的原始日期文本",
            "is_expired": "是否已过期（true/false）",
            "days_until_expiry": "距离过期天数（负数表示已过期）"
        }
    ],
    "is_valid": "整体有效性判断（true/false）",
    "confidence": "识别置信度（0-1之间的浮点数）",
    "notes": "补充说明或注意事项"
}

请确保返回的是有效的JSON格式。
"""
            
            # 调用视觉AI API
            response = self._call_vision_api(prompt, image_input, provider)
            
            # 解析响应
            result = self._parse_date_validation_response(response)
            
            return self.create_success_result(result, "日期验证完成")
            
        except Exception as e:
            self.logger.error(f"日期验证失败: {str(e)}")
            return self.create_error_result(f"日期验证失败: {str(e)}")
    
    def extract_id_card_info(self, image_input: str, provider: Optional[str] = None) -> Dict[str, Any]:
        """
        提取身份证信息（姓名、性别、身份证号）
        ====================================
        
        使用AI视觉模型识别身份证图像中的关键信息。
        只提取必要的信息，不包含其他敏感数据。
        
        Args:
            image_input (str): 身份证图像文件路径或base64编码数据
            provider (str): 指定使用的模型提供方
            
        Returns:
            Dict[str, Any]: 提取结果
                - success (bool): 处理是否成功
                - data (dict): 身份证信息
                    - name (str): 姓名
                    - gender (str): 性别
                    - id_number (str): 身份证号
                    - confidence (float): 识别置信度
                    - card_side (str): 身份证正面/反面
                - message (str): 处理消息
                
        安全注意：
            - 仅提取必要的三项信息
            - 不存储或传输敏感的详细信息
            - 建议使用后立即清除相关数据
        """
        try:
            provider = provider or self.default_provider
            
            # 构建身份证信息提取提示词
            prompt = """
请分析这张身份证图片，提取以下信息：
1. 姓名
2. 性别  
3. 身份证号码

请只提取这三项基本信息，不要包含其他内容。

请按照以下JSON格式返回分析结果：
{
    "card_side": "身份证正面或反面",
    "name": "姓名",
    "gender": "性别",
    "id_number": "身份证号码",
    "confidence": "识别置信度（0-1之间的浮点数）",
    "is_valid_format": "身份证号格式是否有效（true/false）",
    "notes": "补充说明（如果图片不清晰或信息不完整）"
}

注意事项：
- 如果是身份证反面，某些信息可能无法获取，请如实标注
- 如果图片不清晰或无法识别，请在notes中说明
- 请确保返回的是有效的JSON格式

请确保返回的是有效的JSON格式。
"""
            
            # 调用视觉AI API
            response = self._call_vision_api(prompt, image_input, provider)
            
            # 解析响应
            result = self._parse_id_extraction_response(response)
            
            return self.create_success_result(result, "身份证信息提取完成")
            
        except Exception as e:
            self.logger.error(f"身份证信息提取失败: {str(e)}")
            return self.create_error_result(f"身份证信息提取失败: {str(e)}")
    
    def general_ocr(self, image_input: str, provider: Optional[str] = None) -> Dict[str, Any]:
        """
        通用OCR文本识别
        ===============
        
        对图像进行通用的文本识别，返回识别到的所有文本内容。
        
        Args:
            image_input (str): 图像文件路径或base64编码数据
            provider (str): 指定使用的模型提供方
            
        Returns:
            Dict[str, Any]: 识别结果
                - success (bool): 处理是否成功
                - data (dict): OCR结果
                    - text (str): 识别到的文本内容
                    - confidence (float): 识别置信度
                    - language (str): 主要语言
                - message (str): 处理消息
        """
        try:
            provider = provider or self.default_provider
            
            # 构建通用OCR提示词
            prompt = """
请分析这张图片，识别其中的所有文本内容。

请按照以下JSON格式返回分析结果：
{
    "text": "识别到的完整文本内容",
    "language": "主要语言（如：中文、英文等）",
    "confidence": "识别置信度（0-1之间的浮点数）",
    "text_blocks": [
        {
            "content": "文本块内容",
            "position": "在图片中的大致位置（如：左上、中间、右下等）"
        }
    ],
    "notes": "补充说明"
}

请确保返回的是有效的JSON格式。
"""
            
            # 调用视觉AI API
            response = self._call_vision_api(prompt, image_input, provider)
            
            # 解析响应
            result = self._parse_general_ocr_response(response)
            
            return self.create_success_result(result, "OCR识别完成")
            
        except Exception as e:
            self.logger.error(f"OCR识别失败: {str(e)}")
            return self.create_error_result(f"OCR识别失败: {str(e)}")
    
    def _call_vision_api(self, prompt: str, image_input: str, provider: str) -> str:
        """
        调用视觉AI API的核心方法
        ========================
        
        根据指定的模型提供方调用相应的视觉AI接口。
        
        Args:
            prompt (str): 分析提示词
            image_input (str): 图像输入（文件路径或base64数据）
            provider (str): 模型提供方
            
        Returns:
            str: AI模型的响应文本
            
        支持的提供方：
            - qwen: 阿里云百炼Qwen视觉模型
            - doubao: 字节跳动豆包视觉模型
        """
        # 准备图像数据
        image_data = self._prepare_image_data(image_input)
        
        # 根据提供方选择API调用方法
        provider_key = provider.lower()
        if provider_key in ("qwen", "ali", "dashscope"):
            return self._call_qwen_vision_api(prompt, image_data)
        elif provider_key in ("doubao", "ark", "volc", "volcengine"):
            return self._call_doubao_vision_api(prompt, image_data)
        else:
            # 默认使用Qwen
            return self._call_qwen_vision_api(prompt, image_data)
    
    def _call_qwen_vision_api(self, prompt: str, image_data: str) -> str:
        """
        调用Qwen视觉API
        ===============
        
        使用阿里云百炼平台的Qwen视觉模型进行图像分析。
        
        Args:
            prompt (str): 分析提示词
            image_data (str): base64编码的图像数据
            
        Returns:
            str: Qwen模型的响应文本
        """
        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{image_data}"}
                    }
                ]
            }
        ]
        
        completion = self.qwen_client.chat.completions.create(
            model=self.qwen_model,
            messages=cast(Any, messages),
            stream=False,
            temperature=0.1,  # 较低温度确保识别准确性
        )
        
        return completion.choices[0].message.content or ""
    
    def _call_doubao_vision_api(self, prompt: str, image_data: str) -> str:
        """
        调用豆包（方舟）视觉API
        ======================
        
        使用字节跳动方舟平台的豆包视觉模型进行图像分析。
        
        Args:
            prompt (str): 分析提示词
            image_data (str): base64编码的图像数据
            
        Returns:
            str: 豆包模型的响应文本
        """
        ark_api_key = os.getenv("ARK_API_KEY")
        if not ark_api_key:
            raise ValueError("缺少方舟平台API密钥，请设置环境变量 ARK_API_KEY")
        
        ark_client = OpenAI(
            base_url=self.ark_base_url,
            api_key=ark_api_key,
        )
        
        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{image_data}"}
                    }
                ]
            }
        ]
        
        completion = ark_client.chat.completions.create(
            model=self.doubao_model_id,
            messages=cast(Any, messages),
            stream=False,
            temperature=0.1,  # 较低温度确保识别准确性
        )
        
        return completion.choices[0].message.content or ""
    
    def _prepare_image_data(self, image_input: str) -> str:
        """
        准备图像数据
        ============
        
        将输入的图像转换为base64编码格式。
        支持文件路径和base64数据两种输入方式。
        
        Args:
            image_input (str): 图像文件路径或base64编码数据
            
        Returns:
            str: base64编码的图像数据
        """
        # 如果输入已经是base64数据，直接返回
        if self._is_base64_image(image_input):
            return image_input
        
        # 如果是文件路径，读取并编码
        if os.path.isfile(image_input):
            return self._encode_image_file(image_input)
        
        # 如果都不是，抛出错误
        raise ValueError(f"无效的图像输入: {image_input}")
    
    def _is_base64_image(self, data: str) -> bool:
        """
        检查字符串是否为base64编码的图像数据
        """
        try:
            # 检查是否为有效的base64
            if len(data) % 4 == 0:
                base64.b64decode(data, validate=True)
                return True
        except Exception:
            pass
        return False
    
    def _encode_image_file(self, file_path: str) -> str:
        """
        将图像文件编码为base64字符串
        
        Args:
            file_path (str): 图像文件路径
            
        Returns:
            str: base64编码的图像数据
        """
        try:
            # 使用PIL打开图像（支持多种格式）
            with Image.open(file_path) as img:
                # 转换为RGB模式（确保兼容性）
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # 压缩图像以减少数据传输量
                img = self._resize_image_if_needed(img)
                
                # 保存为JPEG格式到内存
                buffer = io.BytesIO()
                img.save(buffer, format='JPEG', quality=85)
                buffer.seek(0)
                
                # 编码为base64
                image_data = base64.b64encode(buffer.read()).decode('utf-8')
                return image_data
                
        except Exception as e:
            raise ValueError(f"无法读取图像文件 {file_path}: {str(e)}")
    
    def _resize_image_if_needed(self, img: Image.Image, max_size: int = 2048) -> Image.Image:
        """
        如果图像过大则调整大小
        
        Args:
            img (Image.Image): PIL图像对象
            max_size (int): 最大尺寸（宽或高）
            
        Returns:
            Image.Image: 调整后的图像
        """
        width, height = img.size
        
        # 如果图像尺寸在限制内，直接返回
        if width <= max_size and height <= max_size:
            return img
        
        # 计算缩放比例
        ratio = min(max_size / width, max_size / height)
        new_width = int(width * ratio)
        new_height = int(height * ratio)
        
        # 使用高质量的重采样算法
        return img.resize((new_width, new_height), Image.Resampling.LANCZOS)
    
    def _parse_date_validation_response(self, response: str) -> Dict[str, Any]:
        """
        解析日期验证响应
        ================
        
        将AI返回的日期验证结果解析为结构化数据。
        """
        try:
            # 尝试JSON解析
            result = json.loads(response)
            
            # 标准化结果格式
            return {
                "document_type": result.get("document_type", "未知文档"),
                "dates_found": self._normalize_dates(result.get("dates_found", [])),
                "is_valid": bool(result.get("is_valid", False)),
                "confidence": float(result.get("confidence", 0.0)),
                "notes": result.get("notes", "")
            }
            
        except json.JSONDecodeError:
            # JSON解析失败时的兜底处理
            return {
                "document_type": "解析失败",
                "dates_found": [],
                "is_valid": False,
                "confidence": 0.0,
                "notes": f"响应解析失败: {response[:200]}..."
            }
    
    def _parse_id_extraction_response(self, response: str) -> Dict[str, Any]:
        """
        解析身份证信息提取响应
        ======================
        
        将AI返回的身份证信息解析为结构化数据。
        """
        try:
            # 尝试JSON解析
            result = json.loads(response)
            
            # 标准化结果格式
            return {
                "card_side": result.get("card_side", "未知"),
                "name": result.get("name", ""),
                "gender": result.get("gender", ""),
                "id_number": result.get("id_number", ""),
                "confidence": float(result.get("confidence", 0.0)),
                "is_valid_format": bool(result.get("is_valid_format", False)),
                "notes": result.get("notes", "")
            }
            
        except json.JSONDecodeError:
            # JSON解析失败时的兜底处理
            return {
                "card_side": "解析失败",
                "name": "",
                "gender": "",
                "id_number": "",
                "confidence": 0.0,
                "is_valid_format": False,
                "notes": f"响应解析失败: {response[:200]}..."
            }
    
    def _parse_general_ocr_response(self, response: str) -> Dict[str, Any]:
        """
        解析通用OCR响应
        ===============
        
        将AI返回的OCR识别结果解析为结构化数据。
        """
        try:
            # 尝试JSON解析
            result = json.loads(response)
            
            # 标准化结果格式
            return {
                "text": result.get("text", ""),
                "language": result.get("language", "未知"),
                "confidence": float(result.get("confidence", 0.0)),
                "text_blocks": result.get("text_blocks", []),
                "notes": result.get("notes", "")
            }
            
        except json.JSONDecodeError:
            # JSON解析失败时的兜底处理
            return {
                "text": response,  # 将原始响应作为文本结果
                "language": "未知",
                "confidence": 0.0,
                "text_blocks": [],
                "notes": "使用原始响应作为文本结果"
            }
    
    def _normalize_dates(self, dates_data: List[Dict]) -> List[Dict[str, Any]]:
        """
        标准化日期数据格式
        ==================
        
        确保所有日期数据都有必要的字段和正确的格式。
        """
        normalized_dates = []
        current_date = datetime.now()
        
        for date_item in dates_data:
            if not isinstance(date_item, dict):
                continue
                
            # 解析日期字符串
            date_str = date_item.get("date", "")
            is_expired = False
            days_until_expiry = 0
            
            if date_str:
                try:
                    # 尝试解析日期
                    parsed_date = self._parse_date_string(date_str)
                    if parsed_date:
                        delta = (parsed_date - current_date).days
                        is_expired = delta < 0
                        days_until_expiry = delta
                except Exception:
                    pass
            
            normalized_dates.append({
                "date": date_str,
                "type": date_item.get("type", "未知"),
                "original_text": date_item.get("original_text", date_str),
                "is_expired": is_expired,
                "days_until_expiry": days_until_expiry
            })
        
        return normalized_dates
    
    def _parse_date_string(self, date_str: str) -> Optional[datetime]:
        """
        解析日期字符串
        ==============
        
        尝试多种日期格式进行解析。
        """
        date_formats = [
            "%Y-%m-%d",
            "%Y/%m/%d",
            "%Y年%m月%d日",
            "%Y.%m.%d",
            "%d/%m/%Y",
            "%d-%m-%Y",
            "%m/%d/%Y",
            "%Y-%m",
            "%Y年%m月"
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        return None


# 便捷函数，用于快速创建和使用OCR Agent
def create_ocr_agent(provider: str = "qwen") -> OCRAgent:
    """
    创建OCR Agent实例的便捷函数
    
    Args:
        provider (str): 默认使用的模型提供方
        
    Returns:
        OCRAgent: OCR代理实例
    """
    return OCRAgent(default_provider=provider)


def validate_document_date(image_path: str, provider: str = "qwen") -> Dict[str, Any]:
    """
    验证文档日期的便捷函数
    
    Args:
        image_path (str): 图像文件路径
        provider (str): 使用的模型提供方
        
    Returns:
        Dict[str, Any]: 验证结果
    """
    agent = create_ocr_agent(provider)
    return agent.validate_document_date(image_path, provider)


def extract_id_card_info(image_path: str, provider: str = "qwen") -> Dict[str, Any]:
    """
    提取身份证信息的便捷函数
    
    Args:
        image_path (str): 身份证图像文件路径
        provider (str): 使用的模型提供方
        
    Returns:
        Dict[str, Any]: 提取结果
    """
    agent = create_ocr_agent(provider)
    return agent.extract_id_card_info(image_path, provider)


# 示例使用
if __name__ == "__main__":
    # 创建OCR代理实例
    ocr_agent = OCRAgent()
    
    # 示例1: 验证证书日期
    # result = ocr_agent.validate_document_date("path/to/certificate.jpg")
    # print("日期验证结果:", result)
    
    # 示例2: 提取身份证信息  
    # result = ocr_agent.extract_id_card_info("path/to/id_card.jpg")
    # print("身份证信息:", result)
    
    # 示例3: 通用OCR识别
    # result = ocr_agent.general_ocr("path/to/document.jpg")
    # print("OCR结果:", result)
    
    print("OCR Agent已就绪，请调用相应的方法进行图像识别。")