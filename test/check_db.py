#!/usr/bin/env python3
"""
检查数据库内容的脚本
"""
import sqlite3
import os

# 检查数据库文件
db_path = 'bid_analysis.db'
if os.path.exists(db_path):
    print(f"数据库文件存在: {db_path}")
    print(f"文件大小: {os.path.getsize(db_path)} bytes")
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 查看所有表
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchall()
        print(f"\n数据库中的表: {[table[0] for table in tables]}")
        
        # 查看files表的记录数量
        if ('files',) in tables:
            cursor.execute("SELECT COUNT(*) FROM files")
            file_count = cursor.fetchone()[0]
            print(f"files表中的记录数量: {file_count}")
            
            # 查看最近的几条记录
            cursor.execute("SELECT id, filename, file_type, upload_time FROM files ORDER BY upload_time DESC LIMIT 5")
            recent_files = cursor.fetchall()
            print("\n最近上传的文件:")
            for file_record in recent_files:
                print(f"  ID: {file_record[0]}")
                print(f"  文件名: {file_record[1]}")
                print(f"  文件类型: {file_record[2]}")
                print(f"  上传时间: {file_record[3]}")
                print("  ---")
        
        # 查看其他表
        for table in tables:
            table_name = table[0]
            if table_name != 'files':
                cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                count = cursor.fetchone()[0]
                print(f"{table_name}表中的记录数量: {count}")
        
        conn.close()
        
    except Exception as e:
        print(f"数据库操作错误: {e}")
else:
    print(f"数据库文件不存在: {db_path}")