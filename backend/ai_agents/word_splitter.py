"""
Word文档拆分器 - 优化版本
专门用于按一级目录拆分Word文档，保留格式和图片
"""

import os
import time
import pythoncom
try:
    import pywintypes  # type: ignore
    _COM_ERROR = pywintypes.com_error
except Exception:  # 兼容编辑器/环境缺少 pywin32 时的类型检查
    class _COM_ERROR(Exception):
        pass
import win32com.client
from .wordtoc_agent import read_word_toc, read_positions_from_markdown

def split_word_document_stable(input_path, output_dir=None, markdown_path=None):
    """
    稳定版本的Word文档拆分器
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"文件不存在: {input_path}")
    
    # 确定输出目录：采用"源文件名称_拆分"的命名规则
    if output_dir is None:
        base_dir = os.path.dirname(input_path)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_dir = os.path.join(base_dir, f"{base_name}_拆分")
    
    os.makedirs(output_dir, exist_ok=True)
    print(f"📂 输出目录: {output_dir}")
    
    # 先尝试使用传入的MD路径，其次回退到默认同名MD
    if not markdown_path:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        markdown_path = os.path.join(os.path.dirname(input_path), f"{base_name}_目录.md")
    print(f"🧭 拆分将尝试使用的目录Markdown: {markdown_path if markdown_path else '无'}")
    
    level1_positions = read_positions_from_markdown(markdown_path)
    
    if level1_positions:
        print("📄 使用已存在的MD文件位置信息...")
        # 从位置信息构建 level1_titles
        level1_titles = []
        for i, pos_info in enumerate(level1_positions):
            level1_titles.append({
                "index": i,
                "title": pos_info["title"],
                "toc_position": i + 1,
                "para_idx": pos_info["para_idx"],
                "start_pos": pos_info["start_pos"]
            })
    else:
        print("🔍 未找到MD文件或位置信息，重新分析文档结构...")
        # 获取目录结构，同时记录位置信息
        toc_items = read_word_toc(input_path, record_positions=True)
        
        # 找到所有一级标题，保存位置信息
        level1_titles = []
        for i, item in enumerate(toc_items):
            if item["层级"] == 1:
                title_info = {
                    "index": i,
                    "title": item["标题文本"],
                    "toc_position": len(level1_titles) + 1
                }
                # 如果有位置信息，保存之
                if "段落索引" in item and "起始位置" in item:
                    title_info["para_idx"] = item["段落索引"]
                    title_info["start_pos"] = item["起始位置"]
                level1_titles.append(title_info)
    
    print(f"📋 识别出 {len(level1_titles)} 个一级标题:")
    for item in level1_titles:
        print(f"  {item['toc_position']}. {item['title']}")
    
    split_files = []
    word_app = None

    def com_retry(callable_obj, *args, _retries=8, _delay=0.4, _backoff=1.6, **kwargs):
        last_exc = None
        for attempt in range(_retries):
            try:
                return callable_obj(*args, **kwargs)
            except _COM_ERROR as e:
                hr = None
                try:
                    if hasattr(e, 'hresult'):
                        hr = e.hresult
                    elif isinstance(e.args, tuple) and len(e.args) > 0:
                        hr = e.args[0]
                except Exception:
                    hr = None
                if hr == -2147418111:  # RPC_E_CALL_REJECTED: 被呼叫方拒绝接收呼叫
                    print("refused，waiting...")
                    time.sleep(_delay)
                    _delay *= _backoff
                    last_exc = e
                    continue
                raise
        if last_exc:
            raise last_exc
        raise RuntimeError("COM 调用失败且未知原因")
    
    abs_input_path = os.path.abspath(input_path)
    try:
        print(f"\n🔄 启动Word应用程序...")
        print(f"📁 文档路径: {abs_input_path}")
        
        # 检查文件是否存在且可读
        if not os.path.exists(abs_input_path):
            raise FileNotFoundError(f"文件不存在: {abs_input_path}")
        
        if not os.access(abs_input_path, os.R_OK):
            raise PermissionError(f"文件无读取权限: {abs_input_path}")
        
        # 检查文件是否被其他程序占用
        try:
            with open(abs_input_path, 'rb') as test_file:
                test_file.read(1)
        except PermissionError:
            raise PermissionError(f"文件被其他程序占用: {abs_input_path}")
        
        # 初始化 COM（单线程），减少 RPC_E_CALL_REJECTED 概率
        try:
            pythoncom.CoInitialize()
            print("✅ COM 初始化成功")
        except Exception as com_init_error:
            print(f"❌ COM 初始化失败: {com_init_error}")
            raise
        
        try:
            word_app = com_retry(win32com.client.Dispatch, "Word.Application")
            print("✅ Word 应用程序创建成功")
        except Exception as word_create_error:
            print(f"❌ Word 应用程序创建失败: {word_create_error}")
            # 尝试替代方案
            try:
                print("🔄 尝试使用 EnsureDispatch...")
                word_app = win32com.client.EnsureDispatch("Word.Application")
                print("✅ Word 应用程序创建成功（EnsureDispatch）")
            except Exception as ensure_error:
                print(f"❌ EnsureDispatch 也失败: {ensure_error}")
                raise word_create_error
        
        word_app.Visible = False
        try:
            word_app.DisplayAlerts = 0  # 禁止弹窗
        except Exception:
            pass
        try:
            word_app.ScreenUpdating = False  # 减少界面刷新开销
        except Exception:
            pass
        
        # 一次性打开原文档
        print("📖 打开原文档...")
        try:
            source_doc = com_retry(word_app.Documents.Open, abs_input_path)
            print("✅ 文档打开成功")
        except Exception as open_error:
            print(f"❌ 文档打开失败: {open_error}")
            # 尝试只读模式打开
            try:
                print("🔄 尝试只读模式打开...")
                source_doc = com_retry(word_app.Documents.Open, abs_input_path, ReadOnly=True)
                print("✅ 文档以只读模式打开成功")
            except Exception as readonly_error:
                print(f"❌ 只读模式也失败: {readonly_error}")
                raise open_error
        
        # 为每个一级标题单独处理
        for i, title_info in enumerate(level1_titles):
            print(f"\n📄 处理章节 {i+1}/{len(level1_titles)}: {title_info['title']}")

            max_attempts = 3
            attempt = 1
            while attempt <= max_attempts:
                new_doc = None
                try:
                    # 创建新文档
                    new_doc = com_retry(word_app.Documents.Add)

                    # 获取当前一级标题信息
                    target_title = title_info['title']
                    next_title = level1_titles[i+1]['title'] if i + 1 < len(level1_titles) else None
                    
                    # 使用预记录的位置信息（如果可用）
                    if 'start_pos' in title_info:
                        start_range = title_info['start_pos']
                        start_para_idx = title_info.get('para_idx')
                        found_start = True
                        print(f"   📍 使用预记录位置: {start_range}")
                        
                        # 计算结束位置
                        if i + 1 < len(level1_titles) and 'start_pos' in level1_titles[i+1]:
                            end_range = level1_titles[i+1]['start_pos'] - 1
                            print(f"   📍 使用下一标题位置: {end_range}")
                        elif i == len(level1_titles) - 1:
                            # 最后一章的优化逻辑保持不变
                            is_empty_after = True
                            try:
                                total_paras = com_retry(lambda: source_doc.Paragraphs.Count)
                            except Exception:
                                total_paras = source_doc.Paragraphs.Count
                            if start_para_idx is not None:
                                max_check = min(total_paras, start_para_idx + 10)
                                for p in range(start_para_idx + 1, max_check + 1):
                                    try:
                                        ptext = source_doc.Paragraphs(p).Range.Text.strip().replace('\r','').replace('\x07','')
                                        if ptext:
                                            is_empty_after = False
                                            break
                                    except Exception:
                                        continue
                            if is_empty_after and start_range is not None:
                                end_range = com_retry(lambda: source_doc.Paragraphs(start_para_idx).Range.End)
                                print("   📍 末章无正文，仅复制标题")
                            else:
                                try:
                                    sec_end = com_retry(lambda: source_doc.Paragraphs(start_para_idx).Range.Sections(1).Range.End)
                                    end_range = sec_end
                                    print("   📍 使用分节边界作为结束位置")
                                except Exception:
                                    end_range = com_retry(lambda: source_doc.Range().End)
                                    print(f"   📍 分节不可用，退回复制到文档末尾")
                        else:
                            end_range = None
                    else:
                        # 回退到原有扫描逻辑（兼容性）
                        found_start = False
                        start_range = None
                        end_range = None
                        start_para_idx = None

                        # 扫描段落找到标题位置
                        try:
                            total_paras = com_retry(lambda: source_doc.Paragraphs.Count)
                        except Exception:
                            total_paras = source_doc.Paragraphs.Count
                        for para_idx in range(1, total_paras + 1):
                            try:
                                para = source_doc.Paragraphs(para_idx)
                                style_name = para.Style.NameLocal
                                para_text = para.Range.Text.strip().replace('\r', '').replace('\x07', '')

                                # 找到起始位置：优先用文本精确匹配
                                if not found_start and para_text == target_title:
                                    start_range = para.Range.Start
                                    start_para_idx = para_idx
                                    found_start = True
                                    print(f"   📍 找到起始位置: 段落 {para_idx}")
                                    continue

                                # 找到结束位置：优先匹配下一个标题文本
                                if found_start and next_title and para_text == next_title:
                                    end_range = para.Range.Start - 1
                                    print(f"   📍 找到结束位置: 段落 {para_idx}")
                                    break

                                # 次选：发现新的一级标题样式（避免误判当前标题自身）
                                if found_start and ("标题 1" in style_name or "Heading 1" in style_name) and para_text != target_title:
                                    end_range = para.Range.Start - 1
                                    print(f"   📍 找到结束位置(样式): 段落 {para_idx}")
                                    break
                            except Exception:
                                continue
                        
                        # 使用 Find 作为备选方案定位下一标题
                        if found_start and end_range is None and next_title:
                            try:
                                end_pos = com_retry(lambda: source_doc.Range().End)
                                search_range = com_retry(source_doc.Range, start_range + 1, end_pos)
                                find = search_range.Find
                                if com_retry(find.Execute, next_title):
                                    end_range = search_range.Start - 1
                                    print(f"   📍 找到结束位置(搜索): 位置 {search_range.Start}")
                            except Exception:
                                pass

                        # 仅当最后一个章节才复制到文档末尾
                        if found_start and end_range is None and (i == len(level1_titles) - 1):
                            # 检查标题后的少量段落是否有正文；若都为空，则仅复制标题行
                            is_empty_after = True
                            try:
                                total_paras = com_retry(lambda: source_doc.Paragraphs.Count)
                            except Exception:
                                total_paras = source_doc.Paragraphs.Count
                            if start_para_idx is not None:
                                max_check = min(total_paras, start_para_idx + 10)
                                for p in range(start_para_idx + 1, max_check + 1):
                                    try:
                                        ptext = source_doc.Paragraphs(p).Range.Text.strip().replace('\r','').replace('\x07','')
                                        if ptext:
                                            is_empty_after = False
                                            break
                                    except Exception:
                                        continue
                            if is_empty_after and start_range is not None:
                                end_range = com_retry(lambda: source_doc.Paragraphs(start_para_idx).Range.End)
                                print("   📍 末章无正文，仅复制标题")
                            else:
                                # 使用分节边界作为末章结束，避免复制至整篇末尾
                                try:
                                    sec_end = com_retry(lambda: source_doc.Paragraphs(start_para_idx).Range.Sections(1).Range.End)
                                    end_range = sec_end
                                    print("   📍 使用分节边界作为结束位置")
                                except Exception:
                                    end_range = com_retry(lambda: source_doc.Range().End)
                                    print(f"   📍 分节不可用，退回复制到文档末尾")
                    
                    if (found_start if 'start_pos' not in title_info else True) and start_range is not None and end_range is not None:
                        try:
                            # 构造要复制的范围
                            copy_range = com_retry(source_doc.Range, start_range, end_range)
                            # 使用 FormattedText，避免系统剪贴板
                            dest_range = com_retry(new_doc.Range, 0, 0)
                            # 读取一次 FormattedText，减少多次COM调用
                            ft = com_retry(lambda: copy_range.FormattedText)
                            def _assign():
                                dest_range.FormattedText = ft
                                return True
                            com_retry(_assign)
                            
                            # 生成文件名（限制长度，规避Windows路径过长）
                            def _sanitize_title(text: str) -> str:
                                allowed = set("-_. ")
                                result = []
                                for ch in text:
                                    if ch.isalnum() or ch in allowed:
                                        result.append(ch)
                                    else:
                                        result.append('_')
                                s = ''.join(result)
                                s = s.replace(' ', '_')
                                while '__' in s:
                                    s = s.replace('__', '_')
                                return s.strip('_')

                            safe_title = _sanitize_title(target_title)
                            if not safe_title:
                                safe_title = f"Chapter_{i+1}"

                            # 动态截断，尽量避免超过 MAX_PATH（260）
                            # 预留目录长度+序号等固定部分，限制标题长度
                            max_total = 240  # 留余量
                            base_len = len(output_dir) + 1 + 2 + 1 + len('.docx')  # dir + '\\' + 'NN' + '_' + ext
                            max_title_len = max(10, max_total - base_len)
                            if len(safe_title) > max_title_len:
                                safe_title = safe_title[:max_title_len]

                            output_filename = f"{i+1:02d}_{safe_title}.docx"
                            output_path = os.path.join(output_dir, output_filename)

                            # 保存文档，失败则回退到短文件名
                            try:
                                com_retry(new_doc.SaveAs2, os.path.abspath(output_path), 16)
                            except Exception as save_err:
                                print(f"   ⚠️ SaveAs2失败，尝试短文件名: {save_err}")
                                short_name = f"{i+1:02d}.docx"
                                output_path = os.path.join(output_dir, short_name)
                                com_retry(new_doc.SaveAs2, os.path.abspath(output_path), 16)
                            
                            # 获取统计信息
                            para_count = new_doc.Paragraphs.Count
                            
                            # 统计图片
                            image_count = 0
                            try:
                                image_count = new_doc.InlineShapes.Count
                            except Exception:
                                pass
                            
                            split_files.append({
                                "path": output_path,
                                "title": target_title,
                                "paragraphs": para_count,
                                "images": image_count,
                                "filename": output_filename
                            })
                            
                            print(f"   ✅ 已保存: {output_filename}")
                            print(f"   📊 段落: {para_count}, 图片: {image_count}")
                            # 成功后跳出重试循环
                            break
                            
                        except Exception as copy_error:
                            print(f"   ❌ 复制粘贴/写入失败: {str(copy_error)}")
                            # 尝试重新建立COM连接
                            try:
                                if new_doc:
                                    com_retry(new_doc.Close, 0)
                                new_doc = com_retry(word_app.Documents.Add)
                            except Exception:
                                pass
                    else:
                        if 'start_pos' in title_info:
                            if end_range is None:
                                print(f"   ❌ 未定位到结束位置，跳过该章节以避免错误")
                        else:
                            if not found_start:
                                print(f"   ❌ 未找到标题位置")
                            elif end_range is None:
                                print(f"   ❌ 未定位到结束位置，跳过该章节以避免错误")
                        # 跳出重试循环，不再重试逻辑错误
                        break

                except Exception as e:
                    hr = None
                    if isinstance(e, _COM_ERROR):
                        try:
                            hr = getattr(e, 'hresult', None) or (e.args[0] if isinstance(e.args, tuple) and e.args else None)
                        except Exception:
                            hr = None
                    print(f"   ❌ 处理失败: {str(e)}")
                    # RPC_E_CALL_REJECTED: 尝试重试
                    if hr in (-2147418111, -2147023174) and attempt < max_attempts:
                        wait_s = min(1.5 * attempt, 3.0)
                        print(f"   🔁 将在 {wait_s:.1f}s 后重试（第 {attempt+1}/{max_attempts} 次）")
                        try:
                            if new_doc:
                                com_retry(new_doc.Close, 0)
                        except Exception:
                            pass
                        # 尝试重新打开源文档以恢复RPC通道
                        try:
                            if 'source_doc' in locals() and source_doc:
                                com_retry(source_doc.Close, 0)
                        except Exception:
                            pass
                        # 若COM服务器不可用，尝试重启Word并重新打开文档
                        try:
                            if hr == -2147023174:
                                try:
                                    if word_app:
                                        com_retry(word_app.Quit)
                                except Exception:
                                    pass
                                # 重建 COM 实例
                                word_app = com_retry(win32com.client.Dispatch, "Word.Application")
                                try:
                                    word_app.DisplayAlerts = 0
                                except Exception:
                                    pass
                                try:
                                    word_app.ScreenUpdating = False
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        # 重新打开原文档
                        try:
                            source_doc = com_retry(word_app.Documents.Open, abs_input_path)
                        except Exception:
                            pass
                        time.sleep(wait_s)
                        attempt += 1
                        continue
                    # 其他异常或最后一次尝试，放弃
                    break
                finally:
                    # 清理新文档资源
                    try:
                        if new_doc:
                            com_retry(new_doc.Close, 0)  # 0 = 不保存
                    except Exception:
                        pass
    
    except Exception as e:
        import traceback as _tb
        print(f"❌ 拆分过程出错: {str(e)}")
        print(f"❌ 错误类型: {type(e).__name__}")
        
        # 特殊处理常见的 COM 错误
        if hasattr(e, 'hresult'):
            print(f"❌ COM 错误代码: {hex(e.hresult) if e.hresult else 'None'}")
            
            # 常见错误代码解释
            error_meanings = {
                -2147418111: "RPC_E_CALL_REJECTED - 应用程序忙碌，拒绝调用",
                -2147024894: "ERROR_FILE_NOT_FOUND - 找不到指定的文件",
                -2147024891: "ERROR_ACCESS_DENIED - 拒绝访问",
                -2147023174: "RPC_E_SERVER_UNAVAILABLE - RPC 服务器不可用",
                -2147352567: "DISP_E_UNKNOWNNAME - 未知名称",
                -2147023144: "ERROR_INVALID_OPERATION - 无效操作"
            }
            
            if e.hresult in error_meanings:
                print(f"❌ 错误含义: {error_meanings[e.hresult]}")
        
        if "Open" in str(e):
            print("💡 这是文档打开相关的错误，可能原因:")
            print("   1. 文件被其他程序占用（如Word正在使用）")
            print("   2. 文件权限不足")
            print("   3. Word应用程序初始化失败")
            print("   4. 文件路径包含特殊字符")
            print("   5. 系统缺少必要的Office组件")
        
        print("🔍 完整堆栈跟踪:")
        print(_tb.format_exc())
    finally:
        # 清理Word应用程序
        try:
            if 'source_doc' in locals():
                com_retry(source_doc.Close, 0)  # 关闭原文档
        except Exception:
            pass
        try:
            if word_app:
                com_retry(word_app.Quit)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    
    # 输出拆分结果摘要
    try:
        print("[RESULT] 拆分结果:")
        print(f"  • 生成文件数: {len(split_files)}")
        for i, info in enumerate(split_files[:10], 1):
            print(f"    {i:02d}. {os.path.basename(info['path'])} | 段落: {info.get('paragraphs')} | 图片: {info.get('images')}")
        if len(split_files) > 10:
            print(f"    ... 还有 {len(split_files) - 10} 个")
    except Exception:
        pass

    return split_files

if __name__ == "__main__":
    # 测试文档
    word_path = r".\技术文件资料.docx"
    
    print("📚 Word文档拆分器 - 稳定版测试")
    print("=" * 60)
    
    try:
        split_result = split_word_document_stable(word_path)
        
        if split_result:
            print(f"\n🎉 拆分完成！共生成 {len(split_result)} 个文件:")
            print("=" * 50)
            
            total_paragraphs = 0
            total_images = 0
            
            for file_info in split_result:
                print(f"📄 {file_info['filename']}")
                print(f"   📋 {file_info['title']}")
                print(f"   📝 段落: {file_info['paragraphs']}, 🖼️ 图片: {file_info['images']}")
                print()
                
                total_paragraphs += file_info['paragraphs']
                total_images += file_info['images']
            
            print(f"📊 总计: {total_paragraphs} 段落, {total_images} 图片")
            print(f"📂 输出目录: {os.path.dirname(split_result[0]['path'])}")
        else:
            print("❌ 拆分失败")
            
    except Exception as e:
        print(f"❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()