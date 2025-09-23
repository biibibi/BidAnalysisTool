#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
数据清理工具 - 快速清理版本
============================

提供快速的数据清理选项，支持命令行参数。

用法：
    python quick_clean.py --help                 # 显示帮助
    python quick_clean.py --status               # 显示当前状态
    python quick_clean.py --tables               # 清空所有数据库表
    python quick_clean.py --temp                 # 清空temp文件夹
    python quick_clean.py --all                  # 清空所有数据和文件
    python quick_clean.py --table files          # 清空指定表
"""

import os
import sys
import sqlite3
import shutil
import argparse
from typing import List

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

try:
    from backend.database import DatabaseManager
except ImportError:
    print("❌ 无法导入DatabaseManager，请确保在项目根目录运行此脚本")
    sys.exit(1)

class QuickCleaner:
    """快速清理工具类"""
    
    def __init__(self):
        self.db_path = os.path.join(project_root, "bid_analysis.db")
        self.temp_path = os.path.join(project_root, "temp")
        self.uploads_path = os.path.join(project_root, "uploads")
        
    def get_table_count(self, table_name: str) -> int:
        """获取表中记录数量"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            count = cursor.fetchone()[0]
            conn.close()
            return count
        except Exception as e:
            print(f"❌ 获取表 {table_name} 记录数失败: {e}")
            return 0
    
    def get_temp_size(self) -> tuple:
        """获取temp文件夹信息"""
        if not os.path.exists(self.temp_path):
            return 0, 0
        
        total_size = 0
        file_count = 0
        
        try:
            for root, dirs, files in os.walk(self.temp_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    try:
                        total_size += os.path.getsize(file_path)
                        file_count += 1
                    except Exception:
                        pass
        except Exception as e:
            print(f"❌ 获取temp文件夹信息失败: {e}")
        
        return file_count, total_size
    
    def get_uploads_size(self) -> tuple:
        """获取uploads文件夹信息"""
        if not os.path.exists(self.uploads_path):
            return 0, 0
        
        total_size = 0
        file_count = 0
        
        try:
            for root, dirs, files in os.walk(self.uploads_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    try:
                        total_size += os.path.getsize(file_path)
                        file_count += 1
                    except Exception:
                        pass
        except Exception as e:
            print(f"❌ 获取uploads文件夹信息失败: {e}")
        
        return file_count, total_size
    
    def format_size(self, size_bytes: int) -> str:
        """格式化文件大小"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
    
    def show_status(self):
        """显示当前数据状态"""
        print("\n📈 当前数据状态：")
        print("-" * 50)
        
        # 数据库表状态
        tables = ['files', 'tender_analysis', 'bid_analysis']
        for table in tables:
            count = self.get_table_count(table)
            print(f"  📊 {table:20} : {count:>6} 条记录")
        
        # temp文件夹状态
        file_count, total_size = self.get_temp_size()
        print(f"  📁 temp文件夹           : {file_count:>6} 个文件 ({self.format_size(total_size)})")
        
        # uploads文件夹状态  
        uploads_count, uploads_size = self.get_uploads_size()
        print(f"  📁 uploads文件夹        : {uploads_count:>6} 个文件 ({self.format_size(uploads_size)})")
        print("-" * 50)
    
    def clear_table(self, table_name: str) -> bool:
        """清空指定数据库表"""
        try:
            print(f"🚀 开始清空表 {table_name}...")
            
            # 检查数据库文件是否存在
            if not os.path.exists(self.db_path):
                print(f"❌ 数据库文件不存在: {self.db_path}")
                return False
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 检查表是否存在
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
            if not cursor.fetchone():
                print(f"❌ 表 {table_name} 不存在")
                conn.close()
                return False
            
            # 获取清空前的记录数
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            before_count = cursor.fetchone()[0]
            print(f"   📊 清空前记录数: {before_count}")
            
            if before_count == 0:
                print(f"   ✅ 表 {table_name} 已经是空的")
                conn.close()
                return True
            
            # 禁用外键约束检查
            cursor.execute("PRAGMA foreign_keys = OFF")
            
            # 清空表
            cursor.execute(f"DELETE FROM {table_name}")
            affected_rows = cursor.rowcount
            
            # 重置自增ID（仅在sqlite_sequence表存在时）
            try:
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='sqlite_sequence'")
                if cursor.fetchone():
                    cursor.execute("DELETE FROM sqlite_sequence WHERE name=?", (table_name,))
            except sqlite3.Error:
                pass  # 忽略sqlite_sequence相关错误
            
            # 重新启用外键约束检查
            cursor.execute("PRAGMA foreign_keys = ON")
            
            # 提交事务
            conn.commit()
            
            # 验证清空结果
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            after_count = cursor.fetchone()[0]
            
            conn.close()
            
            if after_count == 0:
                print(f"   ✅ 成功清空表 {table_name}，删除了 {affected_rows} 条记录")
                return True
            else:
                print(f"   ❌ 清空表 {table_name} 失败，仍有 {after_count} 条记录")
                return False
                
        except sqlite3.Error as e:
            print(f"   ❌ SQLite错误: {e}")
            return False
        except Exception as e:
            print(f"   ❌ 清空表 {table_name} 失败: {e}")
            return False
    
    def clear_all_tables(self) -> bool:
        """清空所有数据库表"""
        print("🚀 清空所有数据库表...")
        
        # 按正确顺序清空（考虑外键约束）
        tables = ['bid_analysis', 'tender_analysis', 'files']
        success = True
        
        for table in tables:
            if not self.clear_table(table):
                success = False
        
        if success:
            print("🎉 所有数据库表已清空")
        else:
            print("⚠️ 部分表清空失败")
        
        return success
    
    def clear_temp_folder(self) -> bool:
        """清空temp文件夹"""
        try:
            if not os.path.exists(self.temp_path):
                print("📁 temp文件夹不存在，无需清理")
                return True
            
            print("🚀 清空temp文件夹...")
            
            # 获取清理前的文件信息
            file_count, total_size = self.get_temp_size()
            print(f"   📊 清理前: {file_count} 个文件 ({self.format_size(total_size)})")
            
            if file_count == 0:
                print("   ✅ temp文件夹已经是空的")
                return True
            
            # 删除temp文件夹内的所有内容
            deleted_count = 0
            for item in os.listdir(self.temp_path):
                item_path = os.path.join(self.temp_path, item)
                try:
                    if os.path.isfile(item_path):
                        os.remove(item_path)
                        deleted_count += 1
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                        deleted_count += 1
                except Exception as e:
                    print(f"   ⚠️ 删除 {item} 失败: {e}")
            
            # 验证清理结果
            final_count, final_size = self.get_temp_size()
            
            if final_count == 0:
                print(f"   ✅ 成功清空temp文件夹，删除了 {deleted_count} 个项目")
                return True
            else:
                print(f"   ⚠️ temp文件夹清理不完整，还剩 {final_count} 个文件")
                return False
            
        except Exception as e:
            print(f"❌ 清空temp文件夹失败: {e}")
            return False
    
    def clear_uploads_folder(self) -> bool:
        """清空uploads文件夹"""
        try:
            if not os.path.exists(self.uploads_path):
                print("📁 uploads文件夹不存在，无需清理")
                return True
            
            print("🚀 清空uploads文件夹...")
            
            # 获取清理前的文件信息
            file_count, total_size = self.get_uploads_size()
            print(f"   📊 清理前: {file_count} 个文件 ({self.format_size(total_size)})")
            
            if file_count == 0:
                print("   ✅ uploads文件夹已经是空的")
                return True
            
            # 删除uploads文件夹内的所有内容
            deleted_count = 0
            for item in os.listdir(self.uploads_path):
                item_path = os.path.join(self.uploads_path, item)
                try:
                    if os.path.isfile(item_path):
                        os.remove(item_path)
                        deleted_count += 1
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                        deleted_count += 1
                except Exception as e:
                    print(f"   ⚠️ 删除 {item} 失败: {e}")
            
            # 验证清理结果
            final_count, final_size = self.get_uploads_size()
            
            if final_count == 0:
                print(f"   ✅ 成功清空uploads文件夹，删除了 {deleted_count} 个项目")
                return True
            else:
                print(f"   ⚠️ uploads文件夹清理不完整，还剩 {final_count} 个文件")
                return False
            
        except Exception as e:
            print(f"❌ 清空uploads文件夹失败: {e}")
            return False
    
    def clear_all(self) -> bool:
        """清空所有数据和文件"""
        print("🚀 开始全面清理...")
        
        success = True
        success &= self.clear_all_tables()
        success &= self.clear_temp_folder()
        success &= self.clear_uploads_folder()
        
        if success:
            print("🎉 全面清理完成！")
        else:
            print("⚠️ 部分清理操作失败")
        
        return success

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="BidAnalysisTool 快速数据清理工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例：
  python quick_clean.py --status               # 显示当前状态
  python quick_clean.py --tables               # 清空所有数据库表
  python quick_clean.py --temp                 # 清空temp文件夹
  python quick_clean.py --all                  # 清空所有数据和文件
  python quick_clean.py --table files          # 清空指定表
        """
    )
    
    parser.add_argument('--status', action='store_true', help='显示当前数据状态')
    parser.add_argument('--tables', action='store_true', help='清空所有数据库表')
    parser.add_argument('--temp', action='store_true', help='清空temp文件夹')
    parser.add_argument('--uploads', action='store_true', help='清空uploads文件夹')
    parser.add_argument('--all', action='store_true', help='清空所有数据和文件')
    parser.add_argument('--table', choices=['files', 'tender_analysis', 'bid_analysis'], 
                       help='清空指定的数据库表')
    
    args = parser.parse_args()
    
    # 如果没有提供任何参数，显示帮助
    if not any(vars(args).values()):
        parser.print_help()
        return
    
    try:
        cleaner = QuickCleaner()
        
        if args.status:
            cleaner.show_status()
        
        if args.table:
            cleaner.clear_table(args.table)
        
        if args.tables:
            cleaner.clear_all_tables()
        
        if args.temp:
            cleaner.clear_temp_folder()
        
        if args.uploads:
            cleaner.clear_uploads_folder()
        
        if args.all:
            cleaner.clear_all()
            
        # 操作完成后显示状态
        if args.tables or args.temp or args.uploads or args.all or args.table:
            print("\n" + "="*50)
            cleaner.show_status()
        
    except Exception as e:
        print(f"❌ 清理工具执行失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()