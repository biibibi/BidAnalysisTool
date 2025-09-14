#!/usr/bin/env python
# -*- coding: utf-8 -*-
#根据excel文件生成网页所需的废标项数据JS代码
import pandas as pd
import json

def process_reference_data():
    """处理ReferenceList数据并生成JavaScript格式"""
    try:
        # 读取Excel文件
        file_path = r'.\ref\ReferenceList.xlsx'
        df = pd.read_excel(file_path, sheet_name=0)
        
        # 类别映射 - 将Excel中的类别映射到现有CSS样式
        category_mapping = {
            '常规检查': {
                'category': '资质要求',
                'categoryClass': 'category-qualification',
                'icon': 'bi-patch-check'
            },
            '技术常规检查': {
                'category': '技术要求', 
                'categoryClass': 'category-technical',
                'icon': 'bi-gear'
            },
            '商务常规检查': {
                'category': '商务要求',
                'categoryClass': 'category-commercial', 
                'icon': 'bi-file-text'
            },
            '技术定制检查': {
                'category': '其他要求',
                'categoryClass': 'category-format',
                'icon': 'bi-search'
            }
        }
        
        violation_items = []
        current_category_info = None
        
        for index, row in df.iterrows():
            # 获取数据，处理NaN值
            serial = str(row['序号']).strip() if pd.notna(row['序号']) else ''
            category = str(row['类别']).strip() if pd.notna(row['类别']) and str(row['类别']).strip() else ''
            title = str(row['检测项目']).strip() if pd.notna(row['检测项目']) else ''
            requirement = str(row['要求说明']).strip() if pd.notna(row['要求说明']) else ''
            
            # 跳过无效行
            if not title or not requirement:
                continue
                
            # 更新当前类别信息
            if category and category in category_mapping:
                current_category_info = category_mapping[category]
            elif not current_category_info:
                # 如果没有类别信息，使用默认值
                current_category_info = {
                    'category': '其他要求',
                    'categoryClass': 'category-format',
                    'icon': 'bi-search'
                }
            
            # 根据检测项目内容智能选择图标
            icon = current_category_info['icon']
            if '授权' in title or '委托' in title:
                icon = 'bi-person-check'
            elif '营业执照' in title:
                icon = 'bi-award'
            elif '身份证' in title:
                icon = 'bi-person-badge'
            elif '承诺' in title:
                icon = 'bi-check2-circle'
            elif '社保' in title:
                icon = 'bi-person-check'
            elif '技术' in title or '产品' in title or '配置' in title:
                icon = 'bi-box'
            elif '报价' in title or '一览表' in title:
                icon = 'bi-currency-dollar'
            elif '工期' in title or '完工' in title:
                icon = 'bi-clock'
            elif '质保' in title or '维保' in title:
                icon = 'bi-tools'
            elif '项目名称' in requirement or '项目编号' in requirement:
                icon = 'bi-tag'
            elif '页眉' in requirement or '页脚' in requirement:
                icon = 'bi-file-earmark'
            elif '搜索' in requirement:
                icon = 'bi-search'
            elif '环保' in requirement or '节能' in requirement:
                icon = 'bi-leaf'
            elif '低价' in requirement:
                icon = 'bi-graph-down'
            
            # 生成项目数据
            item = {
                'id': f'check_{index}_{int(serial) - 1 if serial.isdigit() else index}',
                'title': title,
                'category': current_category_info['category'],
                'categoryClass': current_category_info['categoryClass'],
                'description': f'检查{title}',
                'requirement': requirement,
                'icon': icon
            }
            
            violation_items.append(item)
        
        print(f"成功处理了 {len(violation_items)} 个废标项")
        
        # 生成JavaScript格式的数组
        js_code = "        const violationItems = [\n"
        for i, item in enumerate(violation_items):
            js_code += f"            {{\n"
            js_code += f"                id: '{item['id']}',\n"
            js_code += f"                title: '{item['title']}',\n"
            js_code += f"                category: '{item['category']}',\n"
            js_code += f"                categoryClass: '{item['categoryClass']}',\n"
            js_code += f"                description: '{item['description']}',\n"
            js_code += f"                requirement: '{item['requirement']}',\n"
            js_code += f"                icon: '{item['icon']}'\n"
            js_code += f"            }}"
            if i < len(violation_items) - 1:
                js_code += ","
            js_code += "\n"
        js_code += "        ];"
        
        print("\n生成的JavaScript代码:")
        print(js_code)
        
        # 保存到文件
        with open('violation_items.js', 'w', encoding='utf-8') as f:
            f.write(js_code)
        
        print("\nJavaScript代码已保存到 violation_items.js 文件")
        
        return js_code
        
    except Exception as e:
        print(f"处理数据时发生错误: {e}")
        return None

if __name__ == "__main__":
    process_reference_data()