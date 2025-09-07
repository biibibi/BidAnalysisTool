import sqlite3
import json
import uuid
from datetime import datetime
from typing import Dict, Optional, List
import os

class DatabaseManager:
    """
    数据库管理服务，使用SQLite存储文件和分析结果
    """
    
    def __init__(self, db_path: str = "bid_analysis.db"):
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """
        初始化数据库表结构
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # 创建文件记录表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS files (
                    id TEXT PRIMARY KEY,
                    filename TEXT NOT NULL,
                    file_path TEXT NOT NULL,
                    content TEXT,
                    upload_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    file_size INTEGER,
                    file_type TEXT
                )
            ''')
            
            # 创建招标文件分析结果表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS tender_analysis (
                    id TEXT PRIMARY KEY,
                    file_id TEXT NOT NULL,
                    analysis_result TEXT NOT NULL,
                    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (file_id) REFERENCES files (id)
                )
            ''')
            
            # 创建投标文件分析结果表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS bid_analysis (
                    id TEXT PRIMARY KEY,
                    file_id TEXT NOT NULL,
                    analysis_result TEXT NOT NULL,
                    tender_analysis_id TEXT,
                    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (file_id) REFERENCES files (id),
                    FOREIGN KEY (tender_analysis_id) REFERENCES tender_analysis (id)
                )
            ''')
            
            conn.commit()
    
    def save_file_record(self, file_id: str, filename: str, file_path: str, content: str) -> bool:
        """
        保存文件记录
        
        Args:
            file_id: 文件唯一ID
            filename: 原始文件名
            file_path: 文件存储路径
            content: 文件内容
            
        Returns:
            是否保存成功
        """
        try:
            file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
            file_type = filename.split('.')[-1].lower() if '.' in filename else 'unknown'
            
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO files (id, filename, file_path, content, file_size, file_type)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (file_id, filename, file_path, content, file_size, file_type))
                conn.commit()
            
            return True
        except Exception as e:
            print(f"保存文件记录失败: {e}")
            return False
    
    def get_file_record(self, file_id: str) -> Optional[Dict]:
        """
        获取文件记录
        
        Args:
            file_id: 文件ID
            
        Returns:
            文件记录字典或None
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT * FROM files WHERE id = ?
                ''', (file_id,))
                
                row = cursor.fetchone()
                if row:
                    return dict(row)
                return None
        except Exception as e:
            print(f"获取文件记录失败: {e}")
            return None
    
    def save_tender_analysis(self, file_id: str, analysis_result: Dict) -> str:
        """
        保存招标文件分析结果
        
        Args:
            file_id: 文件ID
            analysis_result: 分析结果字典
            
        Returns:
            分析记录ID
        """
        analysis_id = str(uuid.uuid4())
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO tender_analysis (id, file_id, analysis_result)
                    VALUES (?, ?, ?)
                ''', (analysis_id, file_id, json.dumps(analysis_result, ensure_ascii=False)))
                conn.commit()
            
            return analysis_id
        except Exception as e:
            print(f"保存招标文件分析结果失败: {e}")
            raise e
    
    def save_bid_analysis(self, file_id: str, analysis_result: Dict, tender_analysis_id: Optional[str] = None) -> str:
        """
        保存投标文件分析结果
        
        Args:
            file_id: 文件ID
            analysis_result: 分析结果字典
            tender_analysis_id: 关联的招标文件分析ID（可选）
            
        Returns:
            分析记录ID
        """
        analysis_id = str(uuid.uuid4())
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO bid_analysis (id, file_id, analysis_result, tender_analysis_id)
                    VALUES (?, ?, ?, ?)
                ''', (analysis_id, file_id, json.dumps(analysis_result, ensure_ascii=False), tender_analysis_id))
                conn.commit()
            
            return analysis_id
        except Exception as e:
            print(f"保存投标文件分析结果失败: {e}")
            raise e
    
    def get_tender_analysis(self, analysis_id: str) -> Optional[Dict]:
        """
        获取招标文件分析结果
        
        Args:
            analysis_id: 分析ID
            
        Returns:
            分析结果字典或None
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT * FROM tender_analysis WHERE id = ?
                ''', (analysis_id,))
                
                row = cursor.fetchone()
                if row:
                    result = dict(row)
                    result['analysis_result'] = json.loads(result['analysis_result'])
                    return result
                return None
        except Exception as e:
            print(f"获取招标文件分析结果失败: {e}")
            return None
    
    def get_bid_analysis(self, analysis_id: str) -> Optional[Dict]:
        """
        获取投标文件分析结果
        
        Args:
            analysis_id: 分析ID
            
        Returns:
            分析结果字典或None
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT * FROM bid_analysis WHERE id = ?
                ''', (analysis_id,))
                
                row = cursor.fetchone()
                if row:
                    result = dict(row)
                    result['analysis_result'] = json.loads(result['analysis_result'])
                    return result
                return None
        except Exception as e:
            print(f"获取投标文件分析结果失败: {e}")
            return None
    
    def get_analysis_result(self, analysis_id: str) -> Optional[Dict]:
        """
        获取分析结果（自动判断类型）
        
        Args:
            analysis_id: 分析ID
            
        Returns:
            分析结果字典或None
        """
        # 先尝试获取招标文件分析结果
        result = self.get_tender_analysis(analysis_id)
        if result:
            result['type'] = 'tender'
            return result
        
        # 再尝试获取投标文件分析结果
        result = self.get_bid_analysis(analysis_id)
        if result:
            result['type'] = 'bid'
            return result
        
        return None
    
    def get_file_analysis_history(self, file_id: str) -> List[Dict]:
        """
        获取文件的分析历史
        
        Args:
            file_id: 文件ID
            
        Returns:
            分析历史列表
        """
        history = []
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                
                # 获取招标文件分析历史
                cursor.execute('''
                    SELECT id, created_time, 'tender' as type 
                    FROM tender_analysis 
                    WHERE file_id = ?
                    ORDER BY created_time DESC
                ''', (file_id,))
                
                for row in cursor.fetchall():
                    history.append(dict(row))
                
                # 获取投标文件分析历史
                cursor.execute('''
                    SELECT id, created_time, 'bid' as type 
                    FROM bid_analysis 
                    WHERE file_id = ?
                    ORDER BY created_time DESC
                ''', (file_id,))
                
                for row in cursor.fetchall():
                    history.append(dict(row))
                
                # 按时间排序
                history.sort(key=lambda x: x['created_time'], reverse=True)
                
        except Exception as e:
            print(f"获取分析历史失败: {e}")
        
        return history
    
    def cleanup_old_records(self, days: int = 30):
        """
        清理旧记录
        
        Args:
            days: 保留天数
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # 删除旧的分析记录
                cursor.execute('''
                    DELETE FROM tender_analysis 
                    WHERE created_time < datetime('now', '-{} days')
                '''.format(days))
                
                cursor.execute('''
                    DELETE FROM bid_analysis 
                    WHERE created_time < datetime('now', '-{} days')
                '''.format(days))
                
                # 删除没有关联分析记录的文件记录
                cursor.execute('''
                    DELETE FROM files 
                    WHERE id NOT IN (
                        SELECT DISTINCT file_id FROM tender_analysis
                        UNION
                        SELECT DISTINCT file_id FROM bid_analysis
                    ) AND upload_time < datetime('now', '-{} days')
                '''.format(days))
                
                conn.commit()
                
        except Exception as e:
            print(f"清理旧记录失败: {e}")
