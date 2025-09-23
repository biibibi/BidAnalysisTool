#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Word文档分割器 - 智能Office切换版本"""

import os
import time
import sys
import win32com.client
import pythoncom
from typing import List, Dict, Optional, Any
from pywintypes import com_error as _COM_ERROR

def com_retry(callable_obj, *args, max_retries=3, **kwargs):
    """COM操作重试包装器"""
    for i in range(max_retries):
        try:
            return callable_obj(*args, **kwargs)
        except _COM_ERROR as e:
            if i == max_retries - 1:
                raise
            hr = getattr(e, 'hresult', None)
            if hr == -2147418111:  # RPC_E_CALL_REJECTED
                time.sleep(0.1 * (i + 1))
            else:
                raise
        except Exception as e:
            if i == max_retries - 1:
                raise
            time.sleep(0.1 * (i + 1))

class OfficeManager:
    """智能Office应用管理器"""
    
    def __init__(self):
        self.app = None
        self.app_type = None
        
    def initialize(self):
        """初始化Office应用"""
        try:
            pythoncom.CoInitialize()
            print("✅ COM 初始化成功")
        except Exception as e:
            print(f"❌ COM 初始化失败: {e}")
            raise
        
        # 尝试顺序：WPS -> Word -> EnsureDispatch (优先使用WPS)
        apps_to_try = [
            ("Kwps.Application", "WPS文字"),
            ("Word.Application", "Microsoft Word"),
            ("EnsureDispatch:Word.Application", "Word (EnsureDispatch)")
        ]
        
        for app_id, app_name in apps_to_try:
            try:
                if app_id.startswith("EnsureDispatch:"):
                    self.app = win32com.client.EnsureDispatch(app_id.split(":", 1)[1])
                else:
                    self.app = com_retry(win32com.client.Dispatch, app_id)
                
                self.app_type = app_name
                print(f"✅ {app_name} 应用程序创建成功")
                
                # 配置应用
                try:
                    self.app.DisplayAlerts = 0
                    self.app.ScreenUpdating = False
                except Exception:
                    pass
                    
                return True
                
            except Exception as e:
                print(f"❌ {app_name} 创建失败: {e}")
                continue
        
        raise Exception("所有Office应用程序都无法创建")
    
    def safe_save(self, doc, output_path, format_code=16):
        """安全保存文档 - 支持多种Office应用"""
        try:
            com_retry(doc.SaveAs2, os.path.abspath(output_path), format_code)
            return True
        except Exception as save_err:
            print(f"   ⚠️ {self.app_type} SaveAs2失败: {save_err}")
            
            # 如果当前是Word，尝试切换到WPS
            if "Word" in self.app_type and "WPS" not in self.app_type:
                return self._try_wps_fallback(doc, output_path, format_code)
            
            # 尝试短文件名
            try:
                short_name = os.path.splitext(os.path.basename(output_path))[0][:10] + ".docx"
                short_path = os.path.join(os.path.dirname(output_path), short_name)
                com_retry(doc.SaveAs2, os.path.abspath(short_path), format_code)
                print(f"   ✅ 使用短文件名保存成功: {short_name}")
                return True
            except Exception as short_err:
                print(f"   ❌ 短文件名也失败: {short_err}")
                return False
    
    def _try_wps_fallback(self, doc, output_path, format_code):
        """WPS fallback处理"""
        try:
            print("   🔄 尝试切换到WPS进行保存...")
            
            # 保存当前内容到临时文件
            temp_path = output_path.replace(".docx", "_temp.docx")
            try:
                # 尝试使用不同的格式
                com_retry(doc.SaveAs, os.path.abspath(temp_path))
                
                # 启动WPS重新处理
                wps_app = win32com.client.Dispatch("Kwps.Application")
                wps_app.DisplayAlerts = 0
                wps_doc = wps_app.Documents.Open(os.path.abspath(temp_path))
                
                # 用WPS保存最终文件
                com_retry(wps_doc.SaveAs2, os.path.abspath(output_path), format_code)
                wps_doc.Close(0)
                wps_app.Quit()
                
                # 删除临时文件
                try:
                    os.remove(temp_path)
                except:
                    pass
                
                print("   ✅ WPS fallback保存成功")
                return True
                
            except Exception as wps_err:
                print(f"   ❌ WPS fallback也失败: {wps_err}")
                try:
                    os.remove(temp_path)
                except:
                    pass
                return False
                
        except Exception as fallback_err:
            print(f"   ❌ WPS fallback初始化失败: {fallback_err}")
            return False
    
    def cleanup(self):
        """清理资源"""
        try:
            if self.app:
                com_retry(self.app.Quit)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def split_word_document_enhanced(input_path: str, output_dir: str = None, markdown_path: str = None) -> List[Dict[str, Any]]:
    """增强版文档分割器 - 智能Office切换"""
    
    abs_input_path = os.path.abspath(input_path)
    if not os.path.exists(abs_input_path):
        raise FileNotFoundError(f"输入文件不存在: {abs_input_path}")
    
    # 创建输出目录
    if output_dir is None:
        # 如果没有指定输出目录，使用原来的逻辑（向后兼容）
        base_name = os.path.splitext(os.path.basename(abs_input_path))[0]
        output_dir = os.path.join(os.path.dirname(abs_input_path), f"{base_name}_拆分")
    
    print(f"📂 输出目录: {output_dir}")
    os.makedirs(output_dir, exist_ok=True)
    
    # 初始化Office管理器
    office = OfficeManager()
    split_files = []
    
    try:
        office.initialize()
        
        # 打开文档
        print(f"📖 打开原文档...")
        source_doc = com_retry(office.app.Documents.Open, abs_input_path)
        print("✅ 文档打开成功")
        
        # 获取所有段落和标题信息
        print("🔍 分析文档结构...")
        paragraphs = source_doc.Paragraphs
        print(f"📄 文档包含 {paragraphs.Count} 个段落")
        
        # 识别一级标题（只按主要章节分割）
        titles = []
        for i in range(1, min(paragraphs.Count + 1, 200)):  # 限制分析范围
            try:
                para = paragraphs(i)
                text = para.Range.Text.strip()
                
                if not text or len(text) < 3:  # 提高最小长度要求
                    continue
                
                # 检查是否为一级标题
                style_name = para.Style.NameLocal.lower()
                font_size = para.Range.Font.Size
                
                # 更严格的一级标题识别规则
                is_main_title = (
                    # 样式匹配
                    ("标题 1" in style_name or "heading 1" in style_name) or
                    # 字体大小较大
                    (font_size and font_size >= 16) or
                    # 明确的一级标题模式
                    (text.startswith("第") and ("部分" in text or "章" in text or "节" in text)) or
                    # 常见的文档结构标识
                    any(pattern in text for pattern in [
                        "第一部分", "第二部分", "第三部分", "第四部分", "第五部分", 
                        "第六部分", "第七部分", "第八部分", "第九部分", "第十部分",
                        "第一章", "第二章", "第三章", "第四章", "第五章",
                        "第六章", "第七章", "第八章", "第九章", "第十章"
                    ])
                )
                
                # 排除明显的二级、三级标题
                is_sub_title = (
                    text.count('.') >= 2 or  # 如 "1.1.1"
                    len([c for c in text if c.isdigit()]) >= 3 or  # 多个数字
                    any(sub_pattern in text for sub_pattern in ["1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9."])
                )
                
                if is_main_title and not is_sub_title:
                    titles.append({
                        'title': text,
                        'para_index': i,
                        'font_size': font_size or 12
                    })
                    print(f"   📋 [一级标题 {len(titles)}] {text}")
                    
            except Exception:
                continue
        
        print(f"✅ 识别出 {len(titles)} 个一级标题")
        
        if not titles:
            print("❌ 未找到任何一级标题，无法分割")
            return []
        
        # 分割文档
        for i, title_info in enumerate(titles):
            print(f"\n📄 处理章节 {i+1}/{len(titles)}: {title_info['title']}")
            
            try:
                # 创建新文档
                new_doc = com_retry(office.app.Documents.Add)
                
                # 确定范围
                start_para = title_info['para_index']
                if i < len(titles) - 1:
                    end_para = titles[i + 1]['para_index'] - 1
                else:
                    end_para = paragraphs.Count
                
                print(f"   📍 复制段落 {start_para} 到 {end_para}")
                
                # 复制内容
                start_range = paragraphs(start_para).Range.Start
                end_range = paragraphs(min(end_para, paragraphs.Count)).Range.End
                
                copy_range = source_doc.Range(start_range, end_range)
                dest_range = new_doc.Range(0, 0)
                dest_range.FormattedText = copy_range.FormattedText
                
                # 生成文件名
                safe_title = "".join(c if c.isalnum() or c in " -_." else "_" for c in title_info['title'])
                safe_title = safe_title.replace(" ", "_")[:50]  # 限制长度
                
                if not safe_title:
                    safe_title = f"Chapter_{i+1}"
                
                output_filename = f"{i+1:02d}_{safe_title}.docx"
                output_path = os.path.join(output_dir, output_filename)
                
                # 使用智能保存
                if office.safe_save(new_doc, output_path):
                    # 获取统计信息
                    para_count = new_doc.Paragraphs.Count
                    image_count = 0
                    try:
                        image_count = new_doc.InlineShapes.Count
                    except:
                        pass
                    
                    split_files.append({
                        "path": output_path,
                        "title": title_info['title'],
                        "paragraphs": para_count,
                        "images": image_count,
                        "filename": output_filename
                    })
                    
                    print(f"   ✅ 已保存: {output_filename}")
                    print(f"   📊 段落: {para_count}, 图片: {image_count}")
                else:
                    print(f"   ❌ 保存失败，跳过此章节")
                
                # 关闭新文档
                try:
                    new_doc.Close(0)
                except:
                    pass
                    
            except Exception as e:
                print(f"   ❌ 处理章节失败: {e}")
                try:
                    if 'new_doc' in locals():
                        new_doc.Close(0)
                except:
                    pass
                continue
        
        # 关闭源文档
        try:
            source_doc.Close(0)
        except:
            pass
            
    except Exception as e:
        print(f"❌ 分割过程出错: {e}")
        import traceback
        traceback.print_exc()
    finally:
        office.cleanup()
    
    print(f"\n✅ 分割完成！共生成 {len(split_files)} 个文件")
    return split_files

def split_word_document_stable(input_path: str, output_dir: str = None, markdown_path: str = None) -> List[Dict[str, Any]]:
    """Word文档拆分器 - 主函数（兼容性函数）"""
    return split_word_document_enhanced(input_path, output_dir, markdown_path)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
    else:
        test_file = r"d:\BidAnalysisTool\test\testfile\测试投标文件.docx"
    
    print("🧪 增强版Word文档分割器测试")
    print("=" * 50)
    
    result = split_word_document_enhanced(test_file)
    
    if result:
        print(f"\n🎉 测试成功！生成了 {len(result)} 个文件:")
        for file_info in result:
            print(f"  📄 {file_info['filename']}")
    else:
        print("\n❌ 测试失败")