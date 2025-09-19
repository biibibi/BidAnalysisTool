import os
import json
import re
from docx import Document
import win32com.client
import tempfile

def read_positions_from_markdown(markdown_path):
    """
    从Markdown文件中读取一级标题位置信息
    :param markdown_path: Markdown文件路径
    :return: 位置信息列表或None
    """
    if not os.path.exists(markdown_path):
        return None
    
    try:
        with open(markdown_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 查找JSON代码块
        json_pattern = r'```json\s*\n(.*?)\n```'
        match = re.search(json_pattern, content, re.DOTALL)
        
        if match:
            json_str = match.group(1)
            positions = json.loads(json_str)
            print(f"✅ 从MD文件读取到 {len(positions)} 个一级标题位置信息")
            return positions
        else:
            print("⚠️ MD文件中未找到位置信息")
            return None
            
    except Exception as e:
        print(f"❌ 读取MD文件位置信息失败: {str(e)}")
        return None

def read_word_toc_docx(input_path):
    """
    通过python-docx解析.docx文档标题样式，生成目录结构
    :param input_path: Word文档路径（.docx格式）
    :return: 目录列表，包含{标题文本, 层级}
    """
    doc = Document(input_path)
    toc_items = []

    for paragraph in doc.paragraphs:
        # 检查段落是否应用了标题样式（Heading 1~9）
        if paragraph.style.name.startswith("Heading"):
            try:
                # 提取层级（如"Heading 1" → 1）
                level = int(paragraph.style.name.replace("Heading ", ""))
                # 提取标题文本（去除空字符）
                title = paragraph.text.strip()
                if title:  # 跳过空标题
                    toc_items.append({"标题文本": title, "层级": level})
            except ValueError:
                # 忽略非标准标题样式（如"Heading 1 - 自定义"）
                continue

    return toc_items

def read_word_toc_doc(input_path, record_positions=False):
    """
    通过COM组件解析.doc文档标题样式，生成目录结构
    :param input_path: Word文档路径（.doc格式）
    :param record_positions: 是否记录标题位置信息（用于后续拆分）
    :return: 目录列表，包含{标题文本, 层级, [段落索引, 起始位置]}
    """
    import pythoncom
    import time
    import pywintypes
    
    toc_items = []
    word_app = None
    doc = None
    
    def com_retry(func, *args, max_retries=3, delay=1.0, **kwargs):
        """COM 调用重试机制"""
        for attempt in range(max_retries):
            try:
                return func(*args, **kwargs)
            except pywintypes.com_error as e:
                if hasattr(e, 'hresult'):
                    hr = e.hresult
                    if hr in (-2147418111, -2147023174):  # RPC_E_CALL_REJECTED, RPC_E_SERVER_UNAVAILABLE
                        if attempt < max_retries - 1:
                            print(f"   🔄 COM 调用失败，{delay:.1f}s 后重试... ({attempt + 1}/{max_retries})")
                            time.sleep(delay)
                            delay *= 1.5  # 指数退避
                            continue
                raise
    
    try:
        # 初始化 COM
        pythoncom.CoInitialize()
        print("✅ COM 初始化成功")
        
        # 创建Word应用程序对象（带重试）
        word_app = com_retry(win32com.client.Dispatch, "Word.Application")
        print("✅ Word 应用程序创建成功")
        
        word_app.Visible = False  # 不显示Word窗口
        
        # 设置 Word 选项以减少问题
        try:
            word_app.DisplayAlerts = 0  # 禁用警告对话框
            word_app.ScreenUpdating = False  # 禁用屏幕更新
        except Exception:
            pass
        
        # 打开文档（带重试）
        doc = com_retry(word_app.Documents.Open, os.path.abspath(input_path))
        print(f"✅ 文档打开成功: {os.path.basename(input_path)}")
        
        # 获取段落总数
        try:
            total_paragraphs = com_retry(lambda: doc.Paragraphs.Count)
            print(f"📄 文档包含 {total_paragraphs} 个段落")
        except Exception:
            total_paragraphs = None
        
        # 遍历所有段落，同时记录索引
        para_idx = 1
        while True:
            try:
                # 使用索引访问而不是枚举器，避免 RPC 问题
                paragraph = com_retry(lambda idx=para_idx: doc.Paragraphs(idx))
                
                # 检查段落样式
                style_name = com_retry(lambda: paragraph.Style.NameLocal)
                
                if "标题" in style_name or "Heading" in style_name:
                    try:
                        # 尝试提取层级
                        if "标题" in style_name:
                            # 中文样式：标题 1, 标题 2 等
                            level_str = style_name.replace("标题", "").strip()
                            level = int(level_str) if level_str.isdigit() else 1
                        elif "Heading" in style_name:
                            # 英文样式：Heading 1, Heading 2 等
                            level = int(style_name.replace("Heading", "").strip())
                        else:
                            level = 1
                        
                        # 提取标题文本
                        title = com_retry(lambda: paragraph.Range.Text.strip())
                        # 移除段落标记符
                        title = title.replace('\r', '').replace('\x07', '')
                        
                        if title:  # 跳过空标题
                            item_data = {"标题文本": title, "层级": level}
                            if record_positions:
                                item_data["段落索引"] = para_idx
                                item_data["起始位置"] = com_retry(lambda: paragraph.Range.Start)
                            toc_items.append(item_data)
                            print(f"   📋 [L{level}] {title}")
                    except (ValueError, AttributeError):
                        continue
                
                para_idx += 1
                
                # 如果知道总段落数，可以提前退出
                if total_paragraphs and para_idx > total_paragraphs:
                    break
                    
            except pywintypes.com_error as e:
                if hasattr(e, 'hresult') and e.hresult == -2147024809:  # 超出边界，到达文档末尾
                    break
                else:
                    # 其他 COM 错误，尝试重试
                    try:
                        com_retry(lambda idx=para_idx: doc.Paragraphs(idx), max_retries=1)
                        continue
                    except:
                        print(f"   ⚠️ 段落 {para_idx} 访问失败，跳过")
                        para_idx += 1
                        if para_idx > 1000:  # 防止无限循环
                            break
            except Exception as e:
                print(f"   ⚠️ 段落 {para_idx} 处理异常: {e}")
                para_idx += 1
                if para_idx > 1000:  # 防止无限循环
                    break
                    
    except Exception as e:
        print(f"处理.doc文件时出错: {str(e)}")
        raise
    finally:
        # 关闭文档和Word应用程序
        try:
            if doc:
                com_retry(doc.Close, max_retries=1)
        except Exception:
            pass
        try:
            if word_app:
                com_retry(word_app.Quit, max_retries=1)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    
    print(f"✅ 提取完成，共找到 {len(toc_items)} 个标题")
    return toc_items

def read_word_toc(input_path, record_positions=False):
    """
    自动识别Word文档格式并解析目录结构
    :param input_path: Word文档路径（支持.doc和.docx）
    :param record_positions: 是否记录位置信息（仅对.docx格式通过COM解析有效）
    :return: 目录列表，包含{标题文本, 层级, [段落索引, 起始位置]}
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"文件不存在: {input_path}")
    
    file_ext = os.path.splitext(input_path)[1].lower()
    
    if file_ext == '.docx':
        if record_positions:
            print("检测到.docx格式，使用COM组件解析并记录位置...")
            return read_word_toc_doc(input_path, record_positions)
        else:
            print("检测到.docx格式，使用python-docx解析...")
            return read_word_toc_docx(input_path)
    elif file_ext == '.doc':
        print("检测到.doc格式，使用COM组件解析...")
        return read_word_toc_doc(input_path, record_positions)
    else:
        raise ValueError(f"不支持的文件格式: {file_ext}，仅支持.doc和.docx格式")

def generate_markdown_toc(toc_items, output_path=None, document_title="文档目录", include_positions=False):
    """
    将目录结构生成Markdown格式文件
    :param toc_items: 目录列表，格式为[{标题文本, 层级, [段落索引, 起始位置]}, ...]
    :param output_path: 输出Markdown文件路径，如不提供则自动生成
    :param document_title: 文档标题
    :param include_positions: 是否在MD文件中包含位置信息（用于后续快速拆分）
    :return: Markdown内容字符串
    """
    if not toc_items:
        return "# 目录\n\n暂无目录内容"
    
    # 生成Markdown内容
    markdown_lines = []
    markdown_lines.append(f"# {document_title}")
    markdown_lines.append("")
    markdown_lines.append("---")
    markdown_lines.append("")
    
    # 添加统计信息
    level_counts = {}
    for item in toc_items:
        level = item["层级"]
        level_counts[level] = level_counts.get(level, 0) + 1
    
    markdown_lines.append("## 📊 目录统计")
    markdown_lines.append("")
    markdown_lines.append(f"- **总标题数**: {len(toc_items)}个")
    markdown_lines.append("- **层级分布**:")
    for level in sorted(level_counts.keys()):
        markdown_lines.append(f"  - 层级{level}: {level_counts[level]}个标题")
    
    markdown_lines.append("")
    markdown_lines.append("---")
    markdown_lines.append("")
    
    # 生成目录内容
    markdown_lines.append("## 📋 目录结构")
    markdown_lines.append("")
    
    # 如果包含位置信息，先记录一级标题位置
    if include_positions:
        level1_positions = []
        for i, item in enumerate(toc_items):
            if item["层级"] == 1 and "段落索引" in item and "起始位置" in item:
                level1_positions.append({
                    "title": item["标题文本"],
                    "para_idx": item["段落索引"],
                    "start_pos": item["起始位置"]
                })
        
        if level1_positions:
            markdown_lines.append("### 🎯 一级标题位置信息")
            markdown_lines.append("")
            markdown_lines.append("```json")
            import json
            markdown_lines.append(json.dumps(level1_positions, ensure_ascii=False, indent=2))
            markdown_lines.append("```")
            markdown_lines.append("")
            markdown_lines.append("---")
            markdown_lines.append("")
    
    for i, item in enumerate(toc_items, 1):
        level = item["层级"]
        title = item["标题文本"]
        
        # 使用Markdown标题格式
        if level <= 6:  # Markdown最多支持6级标题
            markdown_title = "#" * level + " " + title
        else:
            # 超过6级的用缩进列表表示
            indent = "  " * (level - 1)
            markdown_title = f"{indent}- {title}"
        
        markdown_lines.append(markdown_title)
    
    # 添加页脚信息
    markdown_lines.append("")
    markdown_lines.append("---")
    markdown_lines.append("")
    markdown_lines.append("*此目录由Word文档目录解析器自动生成*")
    from datetime import datetime
    markdown_lines.append(f"*生成时间: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}*")
    
    markdown_content = "\n".join(markdown_lines)
    
    # 如果提供了输出路径，保存文件
    if output_path:
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            print(f"✅ Markdown目录已保存到: {output_path}")
        except Exception as e:
            print(f"❌ 保存Markdown文件失败: {str(e)}")
    
    return markdown_content

def generate_toc_with_markdown(input_path, output_dir=None):
    """
    解析Word文档并生成Markdown目录文件
    :param input_path: Word文档路径
    :param output_dir: 输出目录，如不提供则使用输入文件同目录
    :return: (目录列表, Markdown文件路径)
    """
    # 解析Word文档目录，记录位置信息以便后续拆分使用
    toc_items = read_word_toc(input_path, record_positions=True)
    
    if not toc_items:
        print("⚠️ 未找到目录内容，无法生成Markdown")
        return toc_items, None
    
    # 确定输出路径
    if output_dir is None:
        output_dir = os.path.dirname(input_path)
    
    # 生成输出文件名：源文件名称_目录.md
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    markdown_filename = f"{base_name}_目录.md"
    output_path = os.path.join(output_dir, markdown_filename)
    
    # 生成文档标题
    document_title = f"{base_name} - 目录结构"
    
    # 生成Markdown内容并保存，包含位置信息
    markdown_content = generate_markdown_toc(toc_items, output_path, document_title, include_positions=True)
    
    return toc_items, output_path

# ------------------- 使用示例 -------------------
if __name__ == "__main__":
    # 测试文档路径
    docx_path = r".\技术文件资料.docx"
    
    print("🔍 Word文档目录解析与Markdown生成测试")
    print("=" * 70)
    
    try:
        print(f"\n📄 正在处理文档: {os.path.basename(docx_path)}")
        
        # 解析文档并生成Markdown
        toc_items, markdown_path = generate_toc_with_markdown(docx_path)
        
        if toc_items and markdown_path:
            print(f"✅ 处理完成！")
            print(f"\n📊 解析结果:")
            print(f"  • 总标题数: {len(toc_items)}个")
            
            # 统计层级
            level_counts = {}
            for item in toc_items:
                level = item["层级"]
                level_counts[level] = level_counts.get(level, 0) + 1
            
            print(f"  • 层级分布:")
            for level in sorted(level_counts.keys()):
                print(f"    - 层级{level}: {level_counts[level]}个")
            
            print(f"\n📝 Markdown文件: {os.path.basename(markdown_path)}")
            
            # 显示前5个标题作为预览
            print(f"\n👀 目录预览（前5项）:")
            print("-" * 40)
            for i, item in enumerate(toc_items[:5], 1):
                indent = "  " * (item["层级"] - 1)
                print(f"{i:2d}. {indent}{item['标题文本']}")
            
            if len(toc_items) > 5:
                print(f"     ... 还有{len(toc_items) - 5}个标题")
                
        else:
            print("❌ 处理失败")
            
    except Exception as e:
        print(f"❌ 处理过程中出现错误：{str(e)}")
    
    print(f"\n{'='*70}")
    print("🎉 测试完成！")
    print("\n📋 功能说明:")
    print("  ✅ 支持.doc和.docx格式")
    print("  ✅ 自动生成Markdown目录文件")
    print("  ✅ 包含统计信息和结构化展示")
    print("  ✅ 保持原文档层级关系")