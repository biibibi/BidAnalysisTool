#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word文档图片分离工具 - 简化版
输入: Word文档 (.docx/.doc)
输出: 分离后的Word文档 + 图片文件夹
"""

import os
import sys
from pathlib import Path
import re

try:
    from docx import Document
    from PIL import Image
    from io import BytesIO
    from openai import OpenAI
except ImportError as e:
    print(f"缺少依赖库: {e}")
    print("请安装: pip install python-docx pillow openai")
    sys.exit(1)

# 可选的.doc文件支持
try:
    import win32com.client
    DOC_SUPPORT = True
except ImportError:
    DOC_SUPPORT = False
    print("提示: 无法处理.doc文件 (需要: pip install pywin32)")


# 初始化火山大模型客户端
def init_ai_client():
    """初始化AI客户端"""
    try:
        client = OpenAI(
            base_url="https://ark.cn-beijing.volces.com/api/v3",
            api_key="5ab3b4ff-bb0f-4289-94a6-1229cfdcfe17",
        )
        return client
    except Exception as e:
        print(f"警告: 无法初始化AI客户端: {e}")
        return None


def generate_image_name(ai_client, context_text, image_index=1, total_images=1):
    """
    使用AI生成图片名称
    
    Args:
        ai_client: AI客户端
        context_text: 图片周围的文字内容
        image_index: 当前图片在连续图片中的索引 
        total_images: 连续图片总数
        
    Returns:
        str: 生成的图片名称（不含扩展名）
    """
    if not ai_client:
        return f"图片_{image_index:03d}"
    
    try:
        # 构建提示词
        if total_images > 1:
            prompt = f"""
基于以下文字内容，为图片生成一个简洁、准确的中文文件名。
这是第{image_index}张图片，共{total_images}张连续图片。

上下文内容：
{context_text[:500]}

要求：
1. 文件名要体现图片内容或相关主题
2. 使用2-8个中文字符
3. 只返回文件名，不要其他内容
4. 如果是连续图片，请用相同的主题词加数字后缀
5. 避免使用特殊字符和空格
"""
        else:
            prompt = f"""
基于以下文字内容，为图片生成一个简洁、准确的中文文件名。

上下文内容：
{context_text[:500]}

要求：
1. 文件名要体现图片内容或相关主题  
2. 使用2-8个中文字符
3. 只返回文件名，不要其他内容
4. 避免使用特殊字符和空格
"""

        completion = ai_client.chat.completions.create(
            model="deepseek-v3-1-250821",
            messages=[
                {"role": "system", "content": "你是一个专业的文件命名助手，擅长根据文档内容生成合适的图片文件名。"},
                {"role": "user", "content": prompt},
            ],
            temperature=0.3,
            max_tokens=50,
            timeout=10  # 添加10秒超时
        )
        
        generated_name = completion.choices[0].message.content.strip()
        
        # 清理文件名，移除特殊字符
        cleaned_name = re.sub(r'[^\w\u4e00-\u9fff]', '', generated_name)
        
        # 如果是连续图片且生成的名称没有数字后缀，添加后缀
        if total_images > 1 and not re.search(r'\d+$', cleaned_name):
            cleaned_name = f"{cleaned_name}_{image_index}"
            
        return cleaned_name if cleaned_name else f"图片_{image_index:03d}"
        
    except Exception as e:
        print(f"AI命名失败: {e}")
        # 简单基于上下文生成描述性名称
        if context_text:
            # 提取关键词
            words = context_text.strip()[:50]
            if "认证" in words:
                return f"认证_{image_index:03d}"
            elif "证书" in words:
                return f"证书_{image_index:03d}"
            elif "测试" in words or "检测" in words:
                return f"测试_{image_index:03d}"
            elif "执照" in words or "营业" in words:
                return f"执照_{image_index:03d}"
            elif "授权" in words:
                return f"授权_{image_index:03d}"
        
        return f"图片_{image_index:03d}"


def is_empty_or_whitespace_paragraph(paragraph):
    """判断段落是否为空或只包含空白符"""
    text = paragraph.text.strip()
    return len(text) == 0


def find_continuous_image_groups(doc, image_parts):
    """
    识别文档中的连续图片组（包括段落和表格中的图片）
    
    Args:
        doc: Word文档对象
        image_parts: 图片部件字典
        
    Returns:
        list: 连续图片组列表，每组包含图片信息和组ID
    """
    all_images = []  # 存储所有图片的位置信息
    
    # 扫描所有段落，记录图片位置
    for para_idx, paragraph in enumerate(doc.paragraphs):
        paragraph_images = []
        
        for run_idx, run in enumerate(paragraph.runs):
            if run._element.xml:
                for drawing in run._element.findall('.//w:drawing', 
                                                   {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    for blip in drawing.findall('.//a:blip', 
                                               {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}):
                        embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        
                        if embed_id and embed_id in image_parts:
                            paragraph_images.append({
                                'para_idx': para_idx,
                                'run_idx': run_idx,
                                'embed_id': embed_id,
                                'paragraph': paragraph,
                                'location_type': 'paragraph',
                                'table_info': None
                            })
        
        all_images.extend(paragraph_images)
    
    # 扫描所有表格，记录图片位置
    table_para_offset = len(doc.paragraphs)  # 表格图片的虚拟段落索引
    
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for cell_para_idx, paragraph in enumerate(cell.paragraphs):
                    for run_idx, run in enumerate(paragraph.runs):
                        if run._element.xml:
                            for drawing in run._element.findall('.//w:drawing', 
                                                               {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                                for blip in drawing.findall('.//a:blip', 
                                                           {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}):
                                    embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                    
                                    if embed_id and embed_id in image_parts:
                                        # 为表格图片分配虚拟段落索引
                                        virtual_para_idx = table_para_offset + table_idx * 100 + row_idx * 10 + cell_idx
                                        all_images.append({
                                            'para_idx': virtual_para_idx,
                                            'run_idx': run_idx,
                                            'embed_id': embed_id,
                                            'paragraph': paragraph,
                                            'location_type': 'table',
                                            'table_info': {
                                                'table_idx': table_idx,
                                                'row_idx': row_idx,
                                                'cell_idx': cell_idx,
                                                'cell_para_idx': cell_para_idx
                                            }
                                        })
    
    # 按位置索引排序
    all_images.sort(key=lambda x: x['para_idx'])
    
    # 识别连续图片组
    image_groups = []
    current_group = []
    group_id = 0
    
    for i, img_info in enumerate(all_images):
        if not current_group:
            # 开始新组
            current_group = [img_info]
        else:
            # 判断是否属于同一连续组
            prev_img = current_group[-1]
            
            # 对于段落图片，使用原有逻辑
            if (img_info['location_type'] == 'paragraph' and 
                prev_img['location_type'] == 'paragraph'):
                
                para_gap = img_info['para_idx'] - prev_img['para_idx']
                is_continuous = True
                
                if para_gap > 0:
                    # 检查中间段落是否都是空白或只有很少文字
                    for check_para_idx in range(prev_img['para_idx'] + 1, img_info['para_idx']):
                        check_para = doc.paragraphs[check_para_idx]
                        if not is_empty_or_whitespace_paragraph(check_para):
                            if len(check_para.text.strip()) > 20:
                                is_continuous = False
                                break
                
                # 如果间隔超过3个段落，也认为不连续
                if para_gap > 3:
                    is_continuous = False
                    
            # 对于表格图片，同一表格同一行的连续单元格认为是连续的
            elif (img_info['location_type'] == 'table' and 
                  prev_img['location_type'] == 'table'):
                
                curr_table = img_info['table_info']
                prev_table = prev_img['table_info']
                
                # 同一表格同一行的相邻单元格
                is_continuous = (curr_table['table_idx'] == prev_table['table_idx'] and
                               curr_table['row_idx'] == prev_table['row_idx'] and
                               abs(curr_table['cell_idx'] - prev_table['cell_idx']) <= 2)
                
            # 表格图片和段落图片之间一般不认为连续
            else:
                is_continuous = False
            
            if is_continuous:
                # 属于当前组
                current_group.append(img_info)
            else:
                # 结束当前组，开始新组
                if len(current_group) > 1:
                    # 只有多于1张图片才算连续组
                    for img in current_group:
                        img['group_id'] = group_id
                        img['group_size'] = len(current_group)
                    group_id += 1
                else:
                    # 单张图片不算组
                    current_group[0]['group_id'] = None
                    current_group[0]['group_size'] = 1
                
                image_groups.extend(current_group)
                current_group = [img_info]
    
    # 处理最后一组
    if current_group:
        if len(current_group) > 1:
            for img in current_group:
                img['group_id'] = group_id
                img['group_size'] = len(current_group)
        else:
            current_group[0]['group_id'] = None
            current_group[0]['group_size'] = 1
        image_groups.extend(current_group)
    
    return image_groups


def convert_doc_to_docx(doc_path):
    """将.doc转换为.docx"""
    if not DOC_SUPPORT:
        raise Exception("不支持.doc格式，请先转换为.docx或安装pywin32")
    
    print(f"正在转换 {doc_path} ...")
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    
    try:
        doc = word_app.Documents.Open(str(Path(doc_path).absolute()))
        docx_path = str(Path(doc_path).with_suffix('.docx'))
        doc.SaveAs2(docx_path, FileFormat=16)
        doc.Close()
        word_app.Quit()
        return docx_path
    except Exception as e:
        word_app.Quit()
        raise e


def extract_and_separate(input_file):
    """
    主要功能：提取图片并分离
    
    Args:
        input_file: 输入的Word文档路径
        
    Returns:
        tuple: (分离后的文档路径, 图片文件夹路径, 图片数量)
    """
    input_path = Path(input_file)
    
    # 检查文件是否存在
    if not input_path.exists():
        raise FileNotFoundError(f"文件不存在: {input_file}")
    
    # 处理.doc文件
    if input_path.suffix.lower() == '.doc':
        input_file = convert_doc_to_docx(input_file)
        input_path = Path(input_file)
    elif input_path.suffix.lower() != '.docx':
        raise ValueError(f"不支持的文件格式: {input_path.suffix}")
    
    # 创建输出路径
    base_name = input_path.stem
    output_doc = input_path.parent / f"{base_name}_分离后.docx"
    output_images = input_path.parent / f"{base_name}_图片"
    output_images.mkdir(exist_ok=True)
    
    # 加载文档
    doc = Document(input_file)
    
    # 初始化AI客户端
    ai_client = init_ai_client()
    
    # 获取所有图片
    image_parts = {}
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image_parts[rel.rId] = rel.target_part

    # 初始化AI客户端
    ai_client = init_ai_client()
    
    # 识别连续图片组
    image_groups = find_continuous_image_groups(doc, image_parts)
    
    image_count = 0
    processed_images = set()  # 记录已处理的图片，避免重复处理

    # 按连续组处理图片
    for img_info in image_groups:
        if img_info['embed_id'] in processed_images:
            continue
            
        image_count += 1
        processed_images.add(img_info['embed_id'])
        
        # 收集上下文
        context_text = ""
        if img_info['location_type'] == 'paragraph':
            # 段落图片：获取前后段落的文字
            para_idx = img_info['para_idx']
            context_range = 2  # 前后各2个段落
            for ctx_idx in range(max(0, para_idx - context_range), 
                               min(len(doc.paragraphs), para_idx + context_range + 1)):
                context_text += doc.paragraphs[ctx_idx].text + " "
        else:
            # 表格图片：获取单元格和相邻单元格的文字
            table_info = img_info['table_info']
            table = doc.tables[table_info['table_idx']]
            row = table.rows[table_info['row_idx']]
            
            # 当前单元格的文字
            cell = row.cells[table_info['cell_idx']]
            for p in cell.paragraphs:
                context_text += p.text + " "
            
            # 相邻单元格的文字
            for adj_cell_idx in range(max(0, table_info['cell_idx'] - 1),
                                    min(len(row.cells), table_info['cell_idx'] + 2)):
                if adj_cell_idx != table_info['cell_idx']:
                    for p in row.cells[adj_cell_idx].paragraphs:
                        context_text += p.text + " "
            
            # 上下行同列单元格的文字（如果存在）
            for adj_row_idx in [table_info['row_idx'] - 1, table_info['row_idx'] + 1]:
                if 0 <= adj_row_idx < len(table.rows):
                    adj_row = table.rows[adj_row_idx]
                    if table_info['cell_idx'] < len(adj_row.cells):
                        for p in adj_row.cells[table_info['cell_idx']].paragraphs:
                            context_text += p.text + " "
        
        # 确定图片在组中的位置
        if img_info['group_id'] is not None:
            # 是连续组的一部分，计算在组中的位置
            group_images = [img for img in image_groups if img['group_id'] == img_info['group_id']]
            group_images.sort(key=lambda x: x['para_idx'])
            img_index_in_group = group_images.index(img_info) + 1
            total_in_group = len(group_images)
        else:
            # 单独图片
            img_index_in_group = 1
            total_in_group = 1
        
        # 生成AI命名
        image_name = generate_image_name(
            ai_client, 
            context_text.strip(),
            img_index_in_group,
            total_in_group
        )
        
        # 保存图片
        image_filename = save_image(
            image_parts[img_info['embed_id']], 
            output_images, 
            image_count,
            custom_name=image_name
        )
        
        # 在文档中替换图片为文本标记
        paragraph = img_info['paragraph']
        run_idx = img_info['run_idx']
        
        # 移除包含图片的run
        if run_idx < len(paragraph.runs):
            paragraph._element.remove(paragraph.runs[run_idx]._element)
        
        # 添加替换文本
        replacement_text = f"[图片: {image_filename}]"
        if paragraph.runs:
            # 如果段落还有其他runs，添加到第一个run
            if paragraph.runs[0].text:
                paragraph.runs[0].text = replacement_text + " " + paragraph.runs[0].text
            else:
                paragraph.runs[0].text = replacement_text
        else:
            # 如果段落没有runs，创建新的run
            paragraph.add_run(replacement_text)

    # 保存分离后的文档
    doc.save(str(output_doc))

    return str(output_doc), str(output_images), image_count


def save_image(image_part, output_dir, counter, custom_name=None):
    """保存图片文件"""
    image_data = image_part.blob
    
    # 检测图片格式
    try:
        img = Image.open(BytesIO(image_data))
        format_name = img.format.lower() if img.format else 'png'
        if format_name == 'jpeg':
            format_name = 'jpg'
    except:
        format_name = 'png'
    
    # 确定文件名
    if custom_name:
        # 使用自定义名称
        filename = f"{custom_name}.{format_name}"
    else:
        # 使用默认命名
        filename = f"图片_{counter:03d}.{format_name}"
    
    # 处理文件名冲突
    filepath = output_dir / filename
    duplicate_counter = 1
    while filepath.exists():
        if custom_name:
            filename = f"{custom_name}_{duplicate_counter}.{format_name}"
        else:
            filename = f"图片_{counter:03d}_{duplicate_counter}.{format_name}"
        filepath = output_dir / filename
        duplicate_counter += 1
    
    # 保存图片
    with open(filepath, 'wb') as f:
        f.write(image_data)
    
    return filename


def main():
    """主函数"""
    print("=" * 50)
    print("Word文档图片分离工具")
    print("=" * 50)
    
    # 检查帮助
    if len(sys.argv) > 1 and sys.argv[1] in ['-h', '--help', 'help']:
        print("使用方法:")
        print(f"  python {Path(__file__).name} <Word文档路径>")
        print("\n示例:")
        print(f"  python {Path(__file__).name} document.docx")
        print(f"  python {Path(__file__).name} '包含空格的文档.docx'")
        print("\n输出:")
        print("  - 文档名_分离后.docx (图文分离后的文档)")
        print("  - 文档名_图片/ (包含所有图片的文件夹)")
        return
    
    # 获取输入文件
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = input("请输入Word文档路径: ").strip('"')
    
    if not input_file:
        print("错误: 未指定输入文件")
        print("使用 --help 查看帮助")
        return
    
    try:
        # 执行分离
        print(f"正在处理: {Path(input_file).name}")
        output_doc, output_images, image_count = extract_and_separate(input_file)
        # 显示结果
        print("\n处理完成!")
        print(f"分离后文档: {Path(output_doc).name}")
        print(f"图片文件夹: {Path(output_images).name}")
        print(f"提取图片数量: {image_count}")
        if image_count == 0:
            print("文档中没有找到图片")
    except Exception as e:
        print(f"处理失败: {e}")


if __name__ == "__main__":
    main()