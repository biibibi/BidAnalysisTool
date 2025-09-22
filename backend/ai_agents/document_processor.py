#!/usr/bin/env python3
"""
文档处理协调器
负责协调目录提取和文档拆分agents的执行
"""

import os
import json
import shutil
from typing import Dict, Any, Optional, Callable
from .wordtoc_agent import read_word_toc, generate_markdown_toc
from .word_splitter import split_word_document_stable


class DocumentProcessor:
    """文档处理协调器类"""
    
    def __init__(self, temp_base_dir: Optional[str] = None):
        """
        初始化文档处理器
        
        Args:
            temp_base_dir: 临时目录基础路径，默认为项目根目录下的temp
        """
        if temp_base_dir is None:
            # 获取项目根目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(os.path.dirname(current_dir))
            temp_base_dir = os.path.join(project_root, "temp")
        
        self.temp_base_dir = temp_base_dir
        os.makedirs(self.temp_base_dir, exist_ok=True)
    
    def process_bid_document(self, file_path: str, file_id: str, progress_callback: Optional[Callable[[int, int, str], None]] = None) -> Dict[str, Any]:
        """
        处理投标文件：生成目录和拆分文档
        
        Args:
            file_path: Word文档路径
            file_id: 文件ID，用于创建子文件夹
            progress_callback: 进度回调函数，接受(current, total, message)参数
            
        Returns:
            处理结果字典
        """
        try:
            # 创建以file_id命名的子文件夹
            work_dir = os.path.join(self.temp_base_dir, file_id)
            os.makedirs(work_dir, exist_ok=True)
            
            if progress_callback:
                progress_callback(10, 100, "开始处理文档...")
            
            # 步骤1: 生成目录结构
            if progress_callback:
                progress_callback(20, 100, "提取文档目录结构...")
            
            toc_result = self._extract_toc(file_path, work_dir, file_id)
            
            if progress_callback:
                progress_callback(50, 100, "目录提取完成，开始拆分文档...")
            
            # 步骤2: 拆分文档
            md_path = toc_result.get('md_path') if isinstance(toc_result.get('md_path'), str) else None
            split_result = self._split_document(file_path, work_dir, file_id, md_path)
            
            if progress_callback:
                progress_callback(90, 100, "文档拆分完成...")
            
            # 整合结果
            result = {
                "success": True,
                "work_dir": work_dir,
                "toc_result": toc_result,
                "split_result": split_result,
                "message": "文档处理完成"
            }
            
            if progress_callback:
                progress_callback(100, 100, "所有处理步骤完成")
            
            return result
            
        except Exception as e:
            error_msg = f"文档处理失败: {str(e)}"
            if progress_callback:
                progress_callback(-1, 100, error_msg)
            
            return {
                "success": False,
                "error": error_msg,
                "work_dir": work_dir if 'work_dir' in locals() else None
            }
    
    def _extract_toc(self, file_path: str, work_dir: str, file_id: str) -> Dict[str, Any]:
        """
        提取文档目录结构
        
        Args:
            file_path: Word文档路径
            work_dir: 工作目录
            file_id: 文件ID
            
        Returns:
            目录提取结果
        """
        try:
            # 生成目录结构
            toc_items = read_word_toc(file_path)
            
            if not toc_items:
                return {
                    "success": False,
                    "error": "无法提取文档目录结构",
                    "toc_items": [],
                    "md_path": None
                }
            
            # 生成markdown文件路径
            md_filename = f"{file_id}_目录.md"
            md_path = os.path.join(work_dir, md_filename)
            
            # 写入markdown文件
            generate_markdown_toc(toc_items, md_path, include_positions=True)
            
            return {
                "success": True,
                "toc_items": toc_items,
                "md_path": md_path,
                "md_filename": md_filename,
                "toc_count": len(toc_items)
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": f"目录提取失败: {str(e)}",
                "toc_items": [],
                "md_path": None
            }
    
    def _split_document(self, file_path: str, work_dir: str, file_id: str, md_path: Optional[str] = None) -> Dict[str, Any]:
        """
        拆分文档
        
        Args:
            file_path: Word文档路径
            work_dir: 工作目录
            file_id: 文件ID
            md_path: 目录markdown文件路径
            
        Returns:
            拆分结果
        """
        try:
            # 创建拆分输出目录
            split_dir = os.path.join(work_dir, "split_documents")
            os.makedirs(split_dir, exist_ok=True)
            
            # 调用文档拆分agent
            split_result = split_word_document_stable(
                input_path=file_path,
                output_dir=split_dir,
                markdown_path=md_path
            )
            
            # 统计拆分结果
            split_files = []
            if os.path.exists(split_dir):
                for item in os.listdir(split_dir):
                    item_path = os.path.join(split_dir, item)
                    if os.path.isfile(item_path) and item.endswith(('.docx', '.doc')):
                        split_files.append({
                            "filename": item,
                            "path": item_path,
                            "size": os.path.getsize(item_path)
                        })
            
            return {
                "success": True,
                "split_dir": split_dir,
                "split_files": split_files,
                "split_count": len(split_files),
                "result": split_result
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": f"文档拆分失败: {str(e)}",
                "split_dir": None,
                "split_files": [],
                "split_count": 0
            }
    
    def get_processing_status(self, file_id: str) -> Dict[str, Any]:
        """
        获取处理状态
        
        Args:
            file_id: 文件ID
            
        Returns:
            处理状态信息
        """
        work_dir = os.path.join(self.temp_base_dir, file_id)
        
        if not os.path.exists(work_dir):
            return {
                "exists": False,
                "work_dir": work_dir
            }
        
        # 检查目录文件
        md_files = [f for f in os.listdir(work_dir) if f.endswith('_目录.md')]
        
        # 检查拆分目录
        split_dir = os.path.join(work_dir, "split_documents")
        split_files = []
        if os.path.exists(split_dir):
            split_files = [f for f in os.listdir(split_dir) if f.endswith(('.docx', '.doc'))]
        
        return {
            "exists": True,
            "work_dir": work_dir,
            "has_toc": len(md_files) > 0,
            "toc_files": md_files,
            "has_splits": len(split_files) > 0,
            "split_files": split_files,
            "split_count": len(split_files)
        }
    
    def cleanup_work_dir(self, file_id: str) -> bool:
        """
        清理工作目录
        
        Args:
            file_id: 文件ID
            
        Returns:
            清理是否成功
        """
        try:
            work_dir = os.path.join(self.temp_base_dir, file_id)
            if os.path.exists(work_dir):
                shutil.rmtree(work_dir)
                return True
            return True
        except Exception as e:
            print(f"清理工作目录失败: {str(e)}")
            return False


# 全局实例
document_processor = DocumentProcessor()