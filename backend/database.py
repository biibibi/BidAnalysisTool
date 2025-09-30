#!/usr/bin/env python3
"""
数据库管理服务模块
==================

本模块提供SQLite数据库的完整管理功能，包括文件记录存储、
分析结果保存、数据查询和维护等核心功能。

数据库设计：
    - files: 文件记录表
    - tender_analysis: 招标文件分析结果表  
    - bid_analysis: 投标文件分析结果表

主要功能：
    1. 数据库初始化和表结构创建
    2. 文件记录的CRUD操作
    3. 分析结果的存储和查询
    4. 历史记录管理
    5. 数据清理和维护

技术特点：
    - 使用SQLite轻量级数据库
    - 支持事务处理
    - 完整的异常处理
    - JSON数据存储支持
    - 外键约束保证数据完整性

依赖库：
    - sqlite3: SQLite数据库接口
    - json: JSON数据处理
    - uuid: 唯一标识生成
    - datetime: 时间处理
    - typing: 类型注解
    - os: 文件系统操作

作者：BidAnalysis Team
创建时间：2025年
版本：1.0
"""

import sqlite3
import json
import uuid
from typing import Dict, Optional, List
import os

class DatabaseManager:
    """
    数据库管理器类
    ==============
    
    负责SQLite数据库的所有操作，包括连接管理、表操作、
    数据存储和查询等功能。
    
    数据表结构：
        files: 文件基础信息
            - id: 文件唯一标识
            - filename: 原始文件名
            - file_path: 存储路径
            - content: 文件文本内容
            - upload_time: 上传时间
            - file_size: 文件大小
            - file_type: 文件类型
            
        tender_analysis: 招标文件分析结果
            - id: 分析记录ID
            - file_id: 关联文件ID
            - analysis_result: 分析结果JSON
            - created_time: 创建时间
            
        bid_analysis: 投标文件分析结果
            - id: 分析记录ID
            - file_id: 关联文件ID
            - analysis_result: 分析结果JSON
            - tender_analysis_id: 关联招标分析ID
            - created_time: 创建时间
    
    使用示例：
        db = DatabaseManager()
        db.save_file_record(file_id, filename, path, content)
        analysis_id = db.save_tender_analysis(file_id, result)
    """
    
    def __init__(self, db_path: Optional[str] = None):
        """
        初始化数据库管理器
        
        Args:
            db_path (str): 数据库文件路径，默认为项目根目录下的bid_analysis.db
            
        功能：
            - 设置数据库文件路径
            - 自动初始化数据库表结构
            - 确保数据库连接正常
        """
        # 如果没有指定路径，使用项目根目录下的数据库文件
        if db_path is None:
            # 获取当前文件（database.py）的目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            # 项目根目录是backend的上级目录
            project_root = os.path.dirname(current_dir)
            # 数据库文件在项目根目录下
            db_path = os.path.join(project_root, "bid_analysis.db")
        
        self.db_path = db_path
        # 初始化数据库表结构
        self.init_database()
    
    def init_database(self):
        """
        初始化数据库表结构
        ==================
        
        创建所有必需的数据表，如果表已存在则跳过。
        建立表之间的外键关系，确保数据完整性。
        
        创建的表：
            1. files - 文件基础信息表
            2. tender_analysis - 招标文件分析结果表
            3. bid_analysis - 投标文件分析结果表
            4. project_info - 项目信息提取结果表
            5. project_matches - 项目信息匹配结果表
            
        表关系：
            - tender_analysis.file_id -> files.id
            - bid_analysis.file_id -> files.id
            - bid_analysis.tender_analysis_id -> tender_analysis.id
            - project_info.file_id -> files.id
            - project_matches.bid_file_id -> files.id
            
        字段说明：
            - 所有主键使用TEXT类型存储UUID
            - 时间字段使用TIMESTAMP类型，默认当前时间
            - JSON数据存储为TEXT类型
            - 外键约束确保数据一致性
            
        异常处理：
            - 数据库锁定
            - 权限不足
            - 磁盘空间不足
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # 创建文件记录表
            # 存储上传文件的基础信息和提取的文本内容
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS files (
                    id TEXT PRIMARY KEY,              -- 文件唯一标识符(UUID)
                    filename TEXT NOT NULL,           -- 原始文件名
                    file_path TEXT NOT NULL,          -- 服务器存储路径
                    content TEXT,                     -- 提取的文本内容
                    upload_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,  -- 上传时间
                    file_size INTEGER,                -- 文件大小(字节)
                    file_type TEXT                    -- 文件类型(扩展名)
                )
            ''')
            
            # 创建招标文件分析结果表
            # 存储AI对招标文件的分析结果
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS tender_analysis (
                    id TEXT PRIMARY KEY,              -- 分析记录唯一标识符
                    file_id TEXT NOT NULL,            -- 关联的文件ID
                    analysis_result TEXT NOT NULL,    -- 分析结果(JSON格式)
                    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,  -- 分析时间
                    project_id TEXT,                  -- 项目编号
                    project_name TEXT,                -- 项目名称
                    FOREIGN KEY (file_id) REFERENCES files (id)  -- 外键约束
                )
            ''')
            
            # 创建投标文件分析结果表
            # 存储AI对投标文件的合规性检查结果
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS bid_analysis (
                    id TEXT PRIMARY KEY,              -- 分析记录唯一标识符
                    file_id TEXT NOT NULL,            -- 关联的投标文件ID
                    analysis_result TEXT NOT NULL,    -- 分析结果(JSON格式)
                    tender_analysis_id TEXT,          -- 关联的招标分析ID(可选)
                    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,  -- 分析时间
                    FOREIGN KEY (file_id) REFERENCES files (id),                    -- 外键约束
                    FOREIGN KEY (tender_analysis_id) REFERENCES tender_analysis (id) -- 外键约束
                )
            ''')

            # 确保新增的项目信息字段存在（兼容旧版数据库）
            self._ensure_column(cursor, 'tender_analysis', 'project_id', 'TEXT')
            self._ensure_column(cursor, 'tender_analysis', 'project_name', 'TEXT')
            
            # 提交事务，确保表创建成功
            conn.commit()

    def _ensure_column(self, cursor: sqlite3.Cursor, table_name: str, column_name: str, column_definition: str):
        """确保指定表包含指定字段，若缺失则自动添加"""
        try:
            cursor.execute(f"PRAGMA table_info({table_name})")
            existing_columns = [row[1] for row in cursor.fetchall()]
            if column_name not in existing_columns:
                cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_definition}")
        except Exception as e:
            print(f"确保表 {table_name} 字段 {column_name} 存在时出错: {e}")
    
    def update_tender_project_info(self, file_id: str, project_id: Optional[str], project_name: Optional[str]) -> bool:
        """更新招标文件分析结果中的项目信息字段"""
        if not file_id:
            return False

        # 如果没有提供任何项目信息，则无需更新
        if project_id is None and project_name is None:
            return False

        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()

                cursor.execute('''
                    UPDATE tender_analysis
                    SET project_id = ?,
                        project_name = ?
                    WHERE file_id = ?
                ''', (project_id, project_name, file_id))

                conn.commit()
                return cursor.rowcount > 0

        except Exception as e:
            print(f"更新招标项目信息失败: {e}")
            return False
    
    def save_file_record(self, file_id: str, filename: str, file_path: str, content: str) -> bool:
        """
        保存文件记录到数据库
        ====================
        
        将上传的文件信息和提取的内容存储到files表中。
        包括文件元数据和文本内容的完整记录。
        
        Args:
            file_id (str): 系统生成的文件唯一标识符
            filename (str): 用户上传的原始文件名
            file_path (str): 文件在服务器上的存储路径
            content (str): 从文件中提取的文本内容
            
        Returns:
            bool: 保存操作结果
                - True: 保存成功
                - False: 保存失败
                
        存储信息：
            - 基础信息：ID、文件名、路径
            - 内容信息：提取的文本内容
            - 元数据：文件大小、类型、上传时间
            
        异常处理：
            - 数据库连接失败
            - 磁盘空间不足
            - 权限问题
            - 数据格式错误
            
        注意：
            - 文件大小获取失败时设为0
            - 文件类型从扩展名提取
            - 上传时间自动设置为当前时间
        """
        try:
            # 安全获取文件大小，文件不存在时设为0
            file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
            
            # 提取文件类型（扩展名），无扩展名时设为'unknown'
            file_type = filename.split('.')[-1].lower() if '.' in filename else 'unknown'
            
            # 使用事务确保数据一致性
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # 插入文件记录
                cursor.execute('''
                    INSERT INTO files (id, filename, file_path, content, file_size, file_type)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (file_id, filename, file_path, content, file_size, file_type))
                
                # 提交事务
                conn.commit()
            
            return True
            
        except Exception as e:
            # 记录错误日志
            print(f"保存文件记录失败: {e}")
            return False
    
    def find_duplicate_file(self, file_size: int) -> Optional[Dict]:
        """
        查找重复文件
        =============
        
        根据文件大小查找是否已存在相同的文件。
        当文件大小完全相同时，认为是同一个文件。
        
        Args:
            file_size (int): 文件大小（字节）
            
        Returns:
            Optional[Dict]: 如果找到重复文件返回文件记录，否则返回None
                - id: 文件ID
                - filename: 文件名
                - file_path: 文件路径
                - file_size: 文件大小
                - upload_time: 上传时间
                - file_type: 文件类型
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                
                # 查找具有相同文件大小的文件
                cursor.execute('''
                    SELECT * FROM files 
                    WHERE file_size = ?
                    ORDER BY upload_time DESC
                    LIMIT 1
                ''', (file_size,))
                
                row = cursor.fetchone()
                if row:
                    result = dict(row)
                    # 验证文件是否确实存在于磁盘上
                    import os
                    if os.path.exists(result['file_path']):
                        return result
                    else:
                        # 如果文件在数据库中存在但磁盘上不存在，应该清理数据库记录
                        self._cleanup_missing_file(result['id'])
                        return None
                
                return None
                
        except Exception as e:
            print(f"查找重复文件失败: {e}")
            return None
    
    def _cleanup_missing_file(self, file_id: str):
        """
        清理数据库中指向不存在文件的记录
        ===============================
        
        当发现数据库中的文件记录指向的文件在磁盘上不存在时，
        清理相关的数据库记录以保持数据一致性。
        
        Args:
            file_id (str): 要清理的文件ID
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # 删除相关的分析记录
                cursor.execute('DELETE FROM tender_analysis WHERE file_id = ?', (file_id,))
                cursor.execute('DELETE FROM bid_analysis WHERE file_id = ?', (file_id,))
                
                # 删除文件记录
                cursor.execute('DELETE FROM files WHERE id = ?', (file_id,))
                
                conn.commit()
                print(f"已清理缺失文件的数据库记录: {file_id}")
                
        except Exception as e:
            print(f"清理缺失文件记录失败: {e}")
    
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


















