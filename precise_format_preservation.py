"""
精确格式保留模块 - 增强PDF转Word的格式保留能力
"""

import os
import traceback
import types

def apply_precise_formatting(converter):
    """
    应用精确格式保留增强功能到PDF转换器
    
    参数:
        converter: PDF转换器实例
    
    返回:
        布尔值，表示是否成功应用增强功能
    """
    print("正在应用精确格式保留增强...")
    
    try:
        # 设置最高格式保留级别
        if hasattr(converter, "format_preservation_level"):
            converter.format_preservation_level = "maximum"
        
        # 启用精确布局保留
        if hasattr(converter, "exact_layout_preservation"):
            converter.exact_layout_preservation = True
        
        # 提高DPI以确保图像质量
        if hasattr(converter, "dpi"):
            converter.dpi = max(int(converter.dpi), 800)
        
        # 启用矢量图形保留
        if hasattr(converter, "preserve_vector_graphics"):
            converter.preserve_vector_graphics = True
        
        # 添加增强的文本块处理方法
        if hasattr(converter, "_process_text_block"):
            original_process_text_block = converter._process_text_block
            
            def enhanced_process_text_block(self, paragraph, block):
                """增强的文本块处理方法"""
                try:
                    # 检查文本块是否有字体信息
                    has_font_info = False
                    font_info = {}
                    
                    # 尝试提取字体信息
                    if "spans" in block:
                        for span in block["spans"]:
                            if "font" in span or "size" in span or "color" in span:
                                has_font_info = True
                                font_info = {
                                    "font": span.get("font", ""),
                                    "size": span.get("size", 11),
                                    "color": span.get("color", (0, 0, 0)),
                                    "bold": "bold" in str(span.get("font", "")).lower(),
                                    "italic": "italic" in str(span.get("font", "")).lower()
                                }
                                break
                    
                    # 如果有字体信息，直接应用到段落
                    if has_font_info:
                        # 清除段落中的现有内容
                        if paragraph.runs:
                            paragraph.clear()
                        
                        # 提取文本
                        text = ""
                        if "text" in block:
                            text = block["text"]
                        elif "spans" in block:
                            for span in block["spans"]:
                                if "text" in span:
                                    text += span["text"]
                        
                        # 创建新的文本运行并应用样式
                        run = paragraph.add_run(text)
                        
                        # 应用字体
                        if "font" in font_info and font_info["font"]:
                            run.font.name = font_info["font"]
                        
                        # 应用字体大小
                        from docx.shared import Pt
                        if "size" in font_info and font_info["size"]:
                            run.font.size = Pt(font_info["size"])
                        
                        # 应用字体样式
                        if "bold" in font_info:
                            run.bold = font_info["bold"]
                        if "italic" in font_info:
                            run.italic = font_info["italic"]
                        
                        # 应用颜色
                        if "color" in font_info and font_info["color"]:
                            from docx.shared import RGBColor
                            r, g, b = font_info["color"]
                            run.font.color.rgb = RGBColor(r, g, b)
                    else:
                        # 如果没有字体信息，使用原始方法
                        original_process_text_block(self, paragraph, block)
                
                except Exception as e:
                    print(f"增强文本块处理时出错: {e}")
                    # 出错时回退到原始方法
                    original_process_text_block(self, paragraph, block)
            
            # 替换原始方法
            converter._process_text_block = types.MethodType(enhanced_process_text_block, converter)
        
        # 增强段落空间处理
        if hasattr(converter, "_process_paragraph"):
            original_process_paragraph = converter._process_paragraph
            
            def enhanced_process_paragraph(self, doc, block, page_num=None):
                """增强的段落处理方法"""
                try:
                    # 调用原始方法处理段落
                    paragraph = original_process_paragraph(self, doc, block, page_num)
                    
                    # 对处理后的段落进行额外的格式优化
                    if paragraph and hasattr(paragraph, "paragraph_format"):
                        # 设置精确的行间距
                        if "line_spacing" in block:
                            from docx.shared import Pt
                            paragraph.paragraph_format.line_spacing = Pt(block["line_spacing"])
                        else:
                            # 使用1.15的默认行间距，接近PDF默认值
                            paragraph.paragraph_format.line_spacing = 1.15
                        
                        # 设置段落前后间距
                        if "space_before" in block:
                            from docx.shared import Pt
                            paragraph.paragraph_format.space_before = Pt(block["space_before"])
                        
                        if "space_after" in block:
                            from docx.shared import Pt
                            paragraph.paragraph_format.space_after = Pt(block["space_after"])
                    
                    return paragraph
                
                except Exception as e:
                    print(f"增强段落处理时出错: {e}")
                    # 出错时回退到原始方法
                    return original_process_paragraph(self, doc, block, page_num)
            
            # 替换原始方法
            converter._process_paragraph = types.MethodType(enhanced_process_paragraph, converter)
        
        # 增强图像处理
        if hasattr(converter, "_process_image_block"):
            original_process_image = converter._process_image_block
            
            def enhanced_process_image(self, doc, pdf_document, page, block):
                """增强的图像处理方法"""
                try:
                    # 提高图像提取质量
                    if hasattr(self, "dpi"):
                        original_dpi = self.dpi
                        self.dpi = max(original_dpi, 1200)  # 临时提高DPI
                    
                    # 调用原始方法
                    result = original_process_image(self, doc, pdf_document, page, block)
                    
                    # 恢复原始DPI
                    if hasattr(self, "dpi") and 'original_dpi' in locals():
                        self.dpi = original_dpi
                    
                    return result
                except Exception as e:
                    print(f"增强图像处理时出错: {e}")
                    # 出错时回退到原始方法
                    return original_process_image(self, doc, pdf_document, page, block)
            
            # 替换原始方法
            converter._process_image_block = types.MethodType(enhanced_process_image, converter)
        
        print("精确格式保留增强功能已成功应用")
        return True
    
    except Exception as e:
        print(f"应用精确格式保留增强时出错: {e}")
        traceback.print_exc()
        return False

# 提供一个验证函数，用于检查格式保留增强是否已应用
def is_precise_formatting_applied(converter):
    """
    检查精确格式保留增强是否已应用
    
    参数:
        converter: PDF转换器实例
    
    返回:
        布尔值，表示是否已应用增强功能
    """
    # 检查关键属性
    format_level_ok = hasattr(converter, "format_preservation_level") and converter.format_preservation_level == "maximum"
    exact_layout_ok = hasattr(converter, "exact_layout_preservation") and converter.exact_layout_preservation
    
    return format_level_ok and exact_layout_ok