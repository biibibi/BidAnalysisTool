#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import json

def read_reference_list():
    """读取ReferenceList.xlsx文件并输出数据"""
    try:
        # 读取Excel文件
        file_path = r'C:\Users\win\Desktop\BidAnalysis\ref\ReferenceList.xlsx'
        
        # 先检查有哪些工作表
        excel_file = pd.ExcelFile(file_path)
        print("Excel文件中的工作表:")
        for sheet_name in excel_file.sheet_names:
            print(f"  - {sheet_name}")
        
        # 读取第一个工作表
        df = pd.read_excel(file_path, sheet_name=0)
        
        print("\n数据预览:")
        print(df.head())
        
        print(f"\n数据形状: {df.shape}")
        print(f"列名: {list(df.columns)}")
        
        # 显示完整数据
        print("\n完整数据:")
        for index, row in df.iterrows():
            print(f"第{index+1}行:")
            for col in df.columns:
                print(f"  {col}: {row[col]}")
            print("-" * 50)
        
        # 将数据转换为JSON格式输出
        print("\nJSON格式:")
        data_list = []
        for index, row in df.iterrows():
            row_dict = {}
            for col in df.columns:
                # 处理NaN值
                value = row[col]
                if pd.isna(value):
                    value = ""
                else:
                    value = str(value).strip()
                row_dict[col] = value
            data_list.append(row_dict)
        
        print(json.dumps(data_list, ensure_ascii=False, indent=2))
        
        return data_list
        
    except Exception as e:
        print(f"读取文件时发生错误: {e}")
        return None

if __name__ == "__main__":
    read_reference_list()