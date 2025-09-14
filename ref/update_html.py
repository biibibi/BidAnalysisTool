#!/usr/bin/env python
# -*- coding: utf-8 -*-

def update_html_violation_items():
    """更新HTML文件中的violationItems数组"""
    
    # 新的violationItems数组
    new_violation_items = '''        // 常见废标项数据 - 基于ReferenceList.xlsx自动生成（按Excel类别分类）
        const violationItems = [
            {
                id: 'check_0_0',
                title: '启用投标专用章的情况说明',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查启用投标专用章的情况说明',
                requirement: '如使用投标专用章，关于启用投标专用章的情况说明是否附上',
                icon: 'bi-patch-check'
            },
            {
                id: 'check_1_1',
                title: '启用投标专用章的情况说明',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查启用投标专用章的情况说明',
                requirement: '项目名称、项目编号等内容是否正确',
                icon: 'bi-tag'
            },
            {
                id: 'check_2_2',
                title: '授权委托书',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查授权委托书',
                requirement: '授权委托书是否附上',
                icon: 'bi-person-check'
            },
            {
                id: 'check_3_3',
                title: '授权委托书',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查授权委托书',
                requirement: '授权人和授权委托人的身份证号码、姓名、项目名称等内容是否正确，内容是否和招标文件要求一致',
                icon: 'bi-person-check'
            },
            {
                id: 'check_4_4',
                title: '法定代表人身份证明',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查法定代表人身份证明',
                requirement: '法定代表人身份证明是否附上',
                icon: 'bi-person-badge'
            },
            {
                id: 'check_5_5',
                title: '法定代表人身份证明',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查法定代表人身份证明',
                requirement: '法定代表人的身份证号码、姓名等内容是否正确，内容是否和招标文件要求一致',
                icon: 'bi-person-badge'
            },
            {
                id: 'check_6_6',
                title: '投标承诺书',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查投标承诺书',
                requirement: '投标承诺书是否附上',
                icon: 'bi-check2-circle'
            },
            {
                id: 'check_7_7',
                title: '投标承诺书',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查投标承诺书',
                requirement: '项目名称、项目编号等内容是否正确，内容是否和招标文件要求一致',
                icon: 'bi-check2-circle'
            },
            {
                id: 'check_8_8',
                title: '授权委托人社保证明',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查授权委托人社保证明',
                requirement: '授权委托人社保是否具备',
                icon: 'bi-person-check'
            },
            {
                id: 'check_9_9',
                title: '授权委托人社保证明',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查授权委托人社保证明',
                requirement: '社保打印日期是否有效',
                icon: 'bi-person-check'
            },
            {
                id: 'check_10_10',
                title: '营业执照',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查营业执照',
                requirement: '总部、省公司、绍兴营业执照是否附上',
                icon: 'bi-award'
            },
            {
                id: 'check_11_11',
                title: '营业执照',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查营业执照',
                requirement: '各营业执照是否在授权的有效期内',
                icon: 'bi-award'
            },
            {
                id: 'check_12_12',
                title: '总公司授权',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查总公司授权',
                requirement: '总公司、省公司授权函是否具备',
                icon: 'bi-person-check'
            },
            {
                id: 'check_13_13',
                title: '总公司授权',
                category: '资质要求',
                categoryClass: 'category-qualification',
                description: '检查总公司授权',
                requirement: '各授权函是否在授权的有效期内',
                icon: 'bi-person-check'
            },
            {
                id: 'check_14_14',
                title: '技术方案（技术偏离表）',
                category: '技术要求',
                categoryClass: 'category-technical',
                description: '检查技术方案（技术偏离表）',
                requirement: '技术偏离表是否附上',
                icon: 'bi-box'
            },
            {
                id: 'check_15_15',
                title: '技术方案（技术偏离表）',
                category: '技术要求',
                categoryClass: 'category-technical',
                description: '检查技术方案（技术偏离表）',
                requirement: '项目名称、项目编号是否正确，表格格式是否和招标文件要求一致',
                icon: 'bi-box'
            },
            {
                id: 'check_16_16',
                title: '投标产品规格配置清单',
                category: '技术要求',
                categoryClass: 'category-technical',
                description: '检查投标产品规格配置清单',
                requirement: '是否有要求实质性响应的项目，以及响应情况检查，是否符合实质性响应的招标要求',
                icon: 'bi-box'
            },
            {
                id: 'check_17_17',
                title: '投标产品规格配置清单',
                category: '技术要求',
                categoryClass: 'category-technical',
                description: '检查投标产品规格配置清单',
                requirement: '投标产品规格配置清单是否具备，填写是否合理，设备数量和表格格式是否和招标文件要求一致',
                icon: 'bi-box'
            },
            {
                id: 'check_18_18',
                title: '投标产品规格配置清单',
                category: '技术要求',
                categoryClass: 'category-technical',
                description: '检查投标产品规格配置清单',
                requirement: '是否有要求实质性响应的项目，如有该项目的具体参数和证明材料是否附上且符合',
                icon: 'bi-box'
            },
            {
                id: 'check_19_19',
                title: '产品检测报告、彩页等',
                category: '技术要求',
                categoryClass: 'category-technical',
                description: '检查产品检测报告、彩页等',
                requirement: '是否有要求实质性响应的项目，如有该项目的检测报告是否具备',
                icon: 'bi-box'
            },
            {
                id: 'check_20_20',
                title: '产品检测报告、彩页等',
                category: '技术要求',
                categoryClass: 'category-technical',
                description: '检查产品检测报告、彩页等',
                requirement: '核心设备的产品彩页等参数证明材料是否具备',
                icon: 'bi-box'
            },
            {
                id: 'check_21_21',
                title: '投标响应函',
                category: '商务要求',
                categoryClass: 'category-commercial',
                description: '检查投标响应函',
                requirement: '投标响应函是否附上，内容是否和招标文件要求一致',
                icon: 'bi-file-text'
            },
            {
                id: 'check_22_22',
                title: '投标响应函',
                category: '商务要求',
                categoryClass: 'category-commercial',
                description: '检查投标响应函',
                requirement: '招标编号、签字人信息等填写内容是否正确',
                icon: 'bi-file-text'
            },
            {
                id: 'check_23_23',
                title: '开标一览表和分项报价表',
                category: '商务要求',
                categoryClass: 'category-commercial',
                description: '检查开标一览表和分项报价表',
                requirement: '商务报价是否超过招标上限价',
                icon: 'bi-currency-dollar'
            },
            {
                id: 'check_24_24',
                title: '开标一览表和分项报价表',
                category: '商务要求',
                categoryClass: 'category-commercial',
                description: '检查开标一览表和分项报价表',
                requirement: '开标一览表分项金额和总价是否一致，总价大小写是否一致',
                icon: 'bi-currency-dollar'
            },
            {
                id: 'check_25_25',
                title: '开标一览表和分项报价表',
                category: '商务要求',
                categoryClass: 'category-commercial',
                description: '检查开标一览表和分项报价表',
                requirement: '分项报价表总价是否和开标一览表一致，具体细项金额检查',
                icon: 'bi-currency-dollar'
            },
            {
                id: 'check_26_26',
                title: '开标一览表和分项报价表',
                category: '商务要求',
                categoryClass: 'category-commercial',
                description: '检查开标一览表和分项报价表',
                requirement: '分项报价表清单内容是否和技术标中投标产品规格配置清单一致',
                icon: 'bi-currency-dollar'
            },
            {
                id: 'check_27_27',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '在招标文件中分别搜索"无效"、"废标"、"▲"、"★"等字样，与投标文件一一对照检查',
                icon: 'bi-search'
            },
            {
                id: 'check_28_28',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '在招标文件中搜索"工期"、"完工"等字样，确认工期，与投标文件一一对照检查',
                icon: 'bi-clock'
            },
            {
                id: 'check_29_29',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '在招标文件中搜索"质保"、"维保"、"维护"等字样，确认项目质保期及维护期等，与投标文件一一对照检查，确认无误请签字',
                icon: 'bi-tools'
            },
            {
                id: 'check_30_30',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '在招标文件中分别搜索"承诺"字样来查找需要承诺的内容，与投标文件一一对照检查',
                icon: 'bi-check2-circle'
            },
            {
                id: 'check_31_31',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '在投标文件中分别搜索"项目名称"、"项目编号"、"招标编号"、"投标单位"等字样，与招标文件中的项目名称与项目编号对照检查',
                icon: 'bi-tag'
            },
            {
                id: 'check_32_32',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '检查投标文件中的页眉及页脚，避免出现其他项目信息',
                icon: 'bi-file-earmark'
            },
            {
                id: 'check_33_33',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '在投标文件中搜索部分单位关键词，避免出现串标情况',
                icon: 'bi-search'
            },
            {
                id: 'check_34_34',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '如有环保与节能产品，节能证书是否已附 节能产品清单：http://www.ccgp.gov.cn/jnhb/jnhbqd/jnqd/',
                icon: 'bi-leaf'
            },
            {
                id: 'check_35_35',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '如低价应标是否已提供低价说明',
                icon: 'bi-graph-down'
            },
            {
                id: 'check_36_36',
                title: '其他项目',
                category: '其他要求',
                categoryClass: 'category-format',
                description: '检查其他项目',
                requirement: '是否出现同一标的物或招标产品(服务)内的主要产品(重要组成部分)出现技术、商务描述不一致或前后描述不一致的情况',
                icon: 'bi-search'
            }
        ];'''
    
    # 读取HTML文件
    html_file_path = r'C:\Users\win\Desktop\BidAnalysis\frontend\bid_analysis.html'
    with open(html_file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 查找并替换violationItems数组
    import re
    
    # 使用正则表达式匹配整个violationItems数组
    pattern = r'([ \t]*)// 常见废标项数据.*?\n([ \t]*)const violationItems = \[.*?\];'
    
    # 使用DOTALL标志让.匹配包括换行符在内的所有字符
    match = re.search(pattern, content, re.DOTALL)
    
    if match:
        # 替换匹配的内容
        new_content = content[:match.start()] + new_violation_items + content[match.end():]
        
        # 写回文件
        with open(html_file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        
        print("成功更新了HTML文件中的violationItems数组")
        print(f"从第{content[:match.start()].count(chr(10)) + 1}行开始替换")
        print(f"替换了{content[match.start():match.end()].count(chr(10)) + 1}行内容")
        
    else:
        print("未找到violationItems数组，尝试手动查找...")
        
        # 查找开始位置
        start_marker = "const violationItems = ["
        start_pos = content.find(start_marker)
        
        if start_pos != -1:
            # 查找结束位置（找到匹配的];）
            bracket_count = 0
            pos = start_pos + len(start_marker)
            end_pos = -1
            
            while pos < len(content):
                if content[pos] == '[':
                    bracket_count += 1
                elif content[pos] == ']':
                    bracket_count -= 1
                    if bracket_count == 0 and pos + 1 < len(content) and content[pos + 1] == ';':
                        end_pos = pos + 2
                        break
                pos += 1
            
            if end_pos != -1:
                # 找到开始行的缩进
                line_start = start_pos
                while line_start > 0 and content[line_start - 1] != '\n':
                    line_start -= 1
                
                indent = content[line_start:start_pos]
                
                # 替换内容
                new_content = content[:line_start] + new_violation_items + content[end_pos:]
                
                with open(html_file_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                
                print("成功手动替换了violationItems数组")
            else:
                print("找到了开始标记但未找到结束标记")
        else:
            print("未找到violationItems数组")

if __name__ == "__main__":
    update_html_violation_items()