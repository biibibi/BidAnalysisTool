"""
Office 应用程序管理器
支持 Microsoft Office, WPS Office, LibreOffice 等
"""

import os
import win32com.client
from typing import Optional, Dict, Any

class OfficeManager:
    """Office 应用程序管理器"""
    
    # 支持的 Office 应用程序配置
    OFFICE_APPS = {
        "word": {
            "name": "Microsoft Word",
            "prog_id": "Word.Application",
            "backup_prog_ids": ["Word.Application"]
        },
        "wps": {
            "name": "WPS 文字",
            "prog_id": "WPS.Application", 
            "backup_prog_ids": ["Kwps.Application", "WPS.Application"]
        },
        "libreoffice": {
            "name": "LibreOffice Writer",
            "prog_id": "com.sun.star.ServiceManager",
            "backup_prog_ids": []
        }
    }
    
    def __init__(self, preferred_app: str = "auto"):
        """
        初始化 Office 管理器
        
        Args:
            preferred_app: 首选应用程序 ("word", "wps", "libreoffice", "auto")
        """
        self.preferred_app = preferred_app
        self.current_app = None
        self.app_instance = None
        
    def get_available_apps(self) -> Dict[str, bool]:
        """检查可用的 Office 应用程序"""
        available = {}
        
        for app_key, app_config in self.OFFICE_APPS.items():
            try:
                # 尝试创建应用程序实例
                app = win32com.client.Dispatch(app_config["prog_id"])
                if app:
                    available[app_key] = True
                    try:
                        app.Quit()
                    except:
                        pass
                else:
                    available[app_key] = False
            except Exception:
                # 尝试备用程序 ID
                found = False
                for backup_id in app_config["backup_prog_ids"]:
                    try:
                        app = win32com.client.Dispatch(backup_id)
                        if app:
                            available[app_key] = True
                            found = True
                            try:
                                app.Quit()
                            except:
                                pass
                            break
                    except Exception:
                        continue
                
                if not found:
                    available[app_key] = False
        
        return available
    
    def create_app_instance(self, app_type: Optional[str] = None) -> Any:
        """
        创建 Office 应用程序实例
        
        Args:
            app_type: 应用程序类型，None 则使用首选或自动检测
            
        Returns:
            Office 应用程序实例
        """
        if app_type is None:
            app_type = self.preferred_app
            
        if app_type == "auto":
            # 自动检测可用应用程序
            available = self.get_available_apps()
            for app_key in ["word", "wps", "libreoffice"]:
                if available.get(app_key, False):
                    app_type = app_key
                    break
            
            if app_type == "auto":
                raise RuntimeError("未找到可用的 Office 应用程序")
        
        if app_type not in self.OFFICE_APPS:
            raise ValueError(f"不支持的应用程序类型: {app_type}")
        
        app_config = self.OFFICE_APPS[app_type]
        last_error = None
        
        # 尝试主程序 ID
        try:
            app = win32com.client.Dispatch(app_config["prog_id"])
            print(f"✅ {app_config['name']} 创建成功")
            self.current_app = app_type
            self.app_instance = app
            return app
        except Exception as e:
            last_error = e
            print(f"❌ {app_config['name']} 主程序 ID 失败: {e}")
        
        # 尝试备用程序 ID
        for backup_id in app_config["backup_prog_ids"]:
            try:
                app = win32com.client.Dispatch(backup_id)
                print(f"✅ {app_config['name']} 创建成功 (备用 ID: {backup_id})")
                self.current_app = app_type
                self.app_instance = app
                return app
            except Exception as e:
                print(f"❌ 备用 ID {backup_id} 失败: {e}")
                last_error = e
        
        raise RuntimeError(f"无法创建 {app_config['name']} 实例: {last_error}")
    
    def get_save_format(self, format_name: str = "docx") -> int:
        """
        获取保存格式代码
        
        Args:
            format_name: 格式名称 ("docx", "doc", "pdf", etc.)
            
        Returns:
            格式代码
        """
        # Microsoft Word / WPS 格式代码
        formats = {
            "docx": 16,  # wdFormatXMLDocument
            "doc": 0,    # wdFormatDocument
            "pdf": 17,   # wdFormatPDF
            "txt": 2,    # wdFormatText
            "rtf": 6     # wdFormatRTF
        }
        
        return formats.get(format_name.lower(), 16)
    
    def safe_save(self, document, file_path: str, format_name: str = "docx", retries: int = 3):
        """
        安全保存文档，支持重试和回退
        
        Args:
            document: 文档对象
            file_path: 保存路径
            format_name: 格式名称
            retries: 重试次数
        """
        format_code = self.get_save_format(format_name)
        abs_path = os.path.abspath(file_path)
        
        for attempt in range(retries):
            try:
                # 尝试 SaveAs2
                if hasattr(document, 'SaveAs2'):
                    document.SaveAs2(abs_path, format_code)
                    return abs_path
                else:
                    # 回退到 SaveAs
                    document.SaveAs(abs_path, format_code)
                    return abs_path
            except Exception as e:
                print(f"❌ 保存失败 (尝试 {attempt + 1}/{retries}): {e}")
                
                if attempt < retries - 1:
                    # 尝试使用短文件名
                    import time
                    time.sleep(0.5)
                    
                    # 生成简化文件名
                    base_dir = os.path.dirname(abs_path)
                    simple_name = f"doc_{int(time.time())}.{format_name}"
                    simple_path = os.path.join(base_dir, simple_name)
                    abs_path = os.path.abspath(simple_path)
                else:
                    raise e
        
        raise RuntimeError(f"保存文档失败，已重试 {retries} 次")
    
    def cleanup(self):
        """清理应用程序实例"""
        if self.app_instance:
            try:
                self.app_instance.Quit()
            except Exception:
                pass
            self.app_instance = None
            self.current_app = None

# 使用示例
if __name__ == "__main__":
    # 检查可用应用程序
    manager = OfficeManager("auto")
    available = manager.get_available_apps()
    print("可用的 Office 应用程序:")
    for app, is_available in available.items():
        status = "✅" if is_available else "❌"
        print(f"  {status} {manager.OFFICE_APPS[app]['name']}")