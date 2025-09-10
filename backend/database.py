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
    
    def __init__(self, db_path: str = "bid_analysis.db"):
        """
        初始化数据库管理器
        
        Args:
            db_path (str): 数据库文件路径，默认为"bid_analysis.db"
            
        功能：
            - 设置数据库文件路径
            - 自动初始化数据库表结构
            - 确保数据库连接正常
        """
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
            
            # 创建项目信息提取结果表
            # 存储从文件中提取的项目编号和项目名称
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS project_info (
                    id TEXT PRIMARY KEY,              -- 信息记录唯一标识符
                    file_id TEXT NOT NULL,            -- 关联的文件ID
                    project_id TEXT,                  -- 提取的项目编号
                    project_name TEXT,                -- 提取的项目名称
                    document_type TEXT,               -- 文档类型(tender/bid)
                    confidence REAL,                  -- 提取置信度(0-1)
                    extraction_data TEXT,             -- 提取详细数据(JSON格式)
                    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,  -- 提取时间
                    FOREIGN KEY (file_id) REFERENCES files (id)  -- 外键约束
                )
            ''')
            
            # 创建项目信息匹配结果表
            # 存储投标文件与招标文件的项目信息匹配结果
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS project_matches (
                    id TEXT PRIMARY KEY,              -- 匹配记录唯一标识符
                    bid_file_id TEXT NOT NULL,        -- 投标文件ID
                    tender_project_id TEXT,           -- 招标项目编号
                    tender_project_name TEXT,         -- 招标项目名称
                    overall_match BOOLEAN,            -- 整体匹配结果
                    match_details TEXT,               -- 匹配详细信息(JSON格式)
                    analysis_report TEXT,             -- 匹配分析报告
                    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,  -- 匹配时间
                    FOREIGN KEY (bid_file_id) REFERENCES files (id)  -- 外键约束
                )
            ''')
            
            # 提交事务，确保表创建成功
            conn.commit()
    
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
                
                # 删除旧的项目信息记录
                cursor.execute('''
                    DELETE FROM project_info 
                    WHERE created_time < datetime('now', '-{} days')
                '''.format(days))
                
                # 删除旧的匹配记录
                cursor.execute('''
                    DELETE FROM project_matches 
                    WHERE created_time < datetime('now', '-{} days')
                '''.format(days))
                
                # 删除没有关联分析记录的文件记录
                cursor.execute('''
                    DELETE FROM files 
                    WHERE id NOT IN (
                        SELECT DISTINCT file_id FROM tender_analysis
                        UNION
                        SELECT DISTINCT file_id FROM bid_analysis
                        UNION
                        SELECT DISTINCT file_id FROM project_info
                    ) AND upload_time < datetime('now', '-{} days')
                '''.format(days))
                
                conn.commit()
                
        except Exception as e:
            print(f"清理旧记录失败: {e}")

    def save_project_info(self, file_id: str, project_data: Dict) -> str:
        """
        保存项目信息提取结果
        ====================
        
        将从文件中提取的项目编号、项目名称等信息保存到数据库。
        
        Args:
            file_id (str): 关联的文件ID
            project_data (Dict): 项目信息数据
                - project_id: 项目编号
                - project_name: 项目名称  
                - document_type: 文档类型
                - confidence: 提取置信度
                - extraction_methods: 提取方法详情
                
        Returns:
            str: 项目信息记录ID
            
        异常处理：
            - 数据库连接失败
            - 参数格式错误
            - 存储空间不足
        """
        info_id = str(uuid.uuid4())
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO project_info (
                        id, file_id, project_id, project_name, document_type, 
                        confidence, extraction_data
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    info_id, 
                    file_id,
                    project_data.get('project_id'),
                    project_data.get('project_name'),
                    project_data.get('document_type'),
                    project_data.get('confidence'),
                    json.dumps(project_data, ensure_ascii=False)
                ))
                conn.commit()
            
            return info_id
            
        except Exception as e:
            print(f"保存项目信息失败: {e}")
            raise e

    def get_project_info(self, info_id: str) -> Optional[Dict]:
        """
        获取项目信息提取结果
        ====================
        
        根据信息ID获取项目信息提取结果。
        
        Args:
            info_id (str): 项目信息记录ID
            
        Returns:
            Optional[Dict]: 项目信息记录或None
                - id: 记录ID
                - file_id: 关联文件ID
                - project_id: 项目编号
                - project_name: 项目名称
                - document_type: 文档类型
                - confidence: 提取置信度
                - extraction_data: 提取详细数据
                - created_time: 创建时间
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT * FROM project_info WHERE id = ?
                ''', (info_id,))
                
                row = cursor.fetchone()
                if row:
                    result = dict(row)
                    result['extraction_data'] = json.loads(result['extraction_data'])
                    return result
                return None
                
        except Exception as e:
            print(f"获取项目信息失败: {e}")
            return None

    def get_project_info_by_file(self, file_id: str) -> Optional[Dict]:
        """
        根据文件ID获取项目信息
        ======================
        
        获取指定文件的最新项目信息提取结果。
        
        Args:
            file_id (str): 文件ID
            
        Returns:
            Optional[Dict]: 最新的项目信息记录或None
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT * FROM project_info 
                    WHERE file_id = ? 
                    ORDER BY created_time DESC 
                    LIMIT 1
                ''', (file_id,))
                
                row = cursor.fetchone()
                if row:
                    result = dict(row)
                    result['extraction_data'] = json.loads(result['extraction_data'])
                    return result
                return None
                
        except Exception as e:
            print(f"根据文件ID获取项目信息失败: {e}")
            return None

    def save_project_match(self, bid_file_id: str, tender_info: Dict, match_result: Dict) -> str:
        """
        保存项目信息匹配结果
        ====================
        
        保存投标文件与招标文件的项目信息匹配结果。
        
        Args:
            bid_file_id (str): 投标文件ID
            tender_info (Dict): 招标文件项目信息
                - project_id: 招标项目编号
                - project_name: 招标项目名称
            match_result (Dict): 匹配结果
                - overall_match: 整体匹配结果
                - project_id_match: 项目编号匹配结果
                - project_name_match: 项目名称匹配结果
                - analysis: 匹配分析报告
                
        Returns:
            str: 匹配记录ID
        """
        match_id = str(uuid.uuid4())
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO project_matches (
                        id, bid_file_id, tender_project_id, tender_project_name,
                        overall_match, match_details, analysis_report
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    match_id,
                    bid_file_id,
                    tender_info.get('project_id'),
                    tender_info.get('project_name'),
                    match_result.get('overall_match'),
                    json.dumps(match_result, ensure_ascii=False),
                    match_result.get('analysis', '')
                ))
                conn.commit()
            
            return match_id
            
        except Exception as e:
            print(f"保存项目匹配结果失败: {e}")
            raise e

    def get_project_match(self, match_id: str) -> Optional[Dict]:
        """
        获取项目信息匹配结果
        ====================
        
        根据匹配ID获取项目信息匹配结果。
        
        Args:
            match_id (str): 匹配记录ID
            
        Returns:
            Optional[Dict]: 匹配记录或None
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT * FROM project_matches WHERE id = ?
                ''', (match_id,))
                
                row = cursor.fetchone()
                if row:
                    result = dict(row)
                    result['match_details'] = json.loads(result['match_details'])
                    return result
                return None
                
        except Exception as e:
            print(f"获取项目匹配结果失败: {e}")
            return None

    def get_project_matches_by_file(self, bid_file_id: str) -> List[Dict]:
        """
        获取文件的所有项目匹配记录
        =========================
        
        获取指定投标文件的所有项目信息匹配记录。
        
        Args:
            bid_file_id (str): 投标文件ID
            
        Returns:
            List[Dict]: 匹配记录列表，按时间倒序排列
        """
        matches = []
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT * FROM project_matches 
                    WHERE bid_file_id = ? 
                    ORDER BY created_time DESC
                ''', (bid_file_id,))
                
                for row in cursor.fetchall():
                    result = dict(row)
                    result['match_details'] = json.loads(result['match_details'])
                    matches.append(result)
                    
        except Exception as e:
            print(f"获取文件匹配记录失败: {e}")
        
        return matches
