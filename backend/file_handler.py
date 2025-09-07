import os
import uuid
from werkzeug.utils import secure_filename
from pathlib import Path
import docx
import PyPDF2
# import fitz  # PyMuPDF - 可选，如果安装失败可注释掉
from typing import Optional

class FileHandler:
    """
    文件处理服务，支持Word和PDF文件的上传和内容提取
    """
    
    ALLOWED_EXTENSIONS = {'.pdf', '.doc', '.docx'}
    
    def __init__(self):
        pass
    
    def is_allowed_file(self, filename: str) -> bool:
        """
        检查文件类型是否被允许
        
        Args:
            filename: 文件名
            
        Returns:
            是否允许上传
        """
        if not filename:
            return False
        
        file_ext = Path(filename).suffix.lower()
        return file_ext in self.ALLOWED_EXTENSIONS
    
    def save_file(self, file, file_id: str, upload_folder: str) -> str:
        """
        保存上传的文件
        
        Args:
            file: 上传的文件对象
            file_id: 文件唯一ID
            upload_folder: 上传目录
            
        Returns:
            保存的文件路径
        """
        # 安全处理文件名
        filename = secure_filename(file.filename)
        file_ext = Path(filename).suffix.lower()
        
        # 生成新的文件名
        new_filename = f"{file_id}{file_ext}"
        file_path = os.path.join(upload_folder, new_filename)
        
        # 保存文件
        file.save(file_path)
        
        return file_path
    
    def extract_content(self, file_path: str) -> str:
        """
        提取文件内容
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件文本内容
        """
        file_ext = Path(file_path).suffix.lower()
        
        try:
            if file_ext == '.pdf':
                return self._extract_pdf_content(file_path)
            elif file_ext in ['.doc', '.docx']:
                return self._extract_word_content(file_path)
            else:
                raise ValueError(f"不支持的文件类型: {file_ext}")
        except Exception as e:
            raise Exception(f"提取文件内容失败: {str(e)}")
    
    def _extract_pdf_content(self, file_path: str) -> str:
        """
        提取PDF文件内容
        """
        content = ""
        
        try:
            # 使用PyPDF2提取PDF内容
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        content += page_text + "\n"
        except Exception as e:
            raise Exception(f"PDF内容提取失败: {str(e)}")
        
        return content.strip()
    
    def _extract_word_content(self, file_path: str) -> str:
        """
        提取Word文件内容
        """
        try:
            # 使用python-docx处理.docx文件
            if file_path.lower().endswith('.docx'):
                doc = docx.Document(file_path)
                content = ""
                
                # 提取段落文本
                for paragraph in doc.paragraphs:
                    content += paragraph.text + "\n"
                
                # 提取表格文本
                for table in doc.tables:
                    for row in table.rows:
                        row_text = []
                        for cell in row.cells:
                            row_text.append(cell.text.strip())
                        content += " | ".join(row_text) + "\n"
                
                return content.strip()
            
            else:
                # 对于.doc文件，需要其他库或转换工具
                raise Exception("暂不支持.doc格式，请转换为.docx格式后上传")
                
        except Exception as e:
            raise Exception(f"Word文档内容提取失败: {str(e)}")
    
    def get_file_info(self, file_path: str) -> dict:
        """
        获取文件基本信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件信息字典
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        file_stat = os.stat(file_path)
        file_ext = Path(file_path).suffix.lower()
        
        return {
            'size': file_stat.st_size,
            'extension': file_ext,
            'modified_time': file_stat.st_mtime,
            'created_time': file_stat.st_ctime
        }
    
    def delete_file(self, file_path: str) -> bool:
        """
        删除文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否删除成功
        """
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        except Exception:
            return False
