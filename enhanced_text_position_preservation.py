"""
增强字体样式和文本位置保留模块
用于解决PDF转Word时字体样式和文本位置保留不准确的问题
"""

import os
import sys
import traceback
import types
import re

def apply_text_position_preservation(converter):
    """
    应用文本位置和字体样式保留增强
    
    参数:
        converter: PDF转换器实例
    
    返回:
        修改后的转换器实例
    """
    try:
        print("应用文本位置和字体样式保留增强...")
        
        # 1. 增强文本块处理方法
        enhance_text_block_processing(converter)
        
        # 2. 增强段落格式检测方法
        enhance_paragraph_format_detection(converter)
        
        # 3. 增强字体映射方法
        enhance_font_mapping(converter)
        
        # 4. 增强文本位置识别
        enhance_text_position_detection(converter)
        
        print("文本位置和字体样式保留增强应用完成")
        return converter
        
    except Exception as e:
        print(f"应用文本位置和字体样式保留增强时出错: {e}")
        traceback.print_exc()
        return converter

def enhance_text_block_processing(converter):
    """增强文本块处理，更精确保留字体样式和格式"""
    try:
        # 检查并增强文本块处理方法
        if hasattr(converter, '_process_text_block'):
            original_process_text_block = converter._process_text_block
            
            def enhanced_process_text_block(self, paragraph, block):
                """增强的文本块处理方法，精确保留字体样式和格式"""
                try:
                    # 提取文本块中的span列表
                    spans = []
                    if "spans" in block:
                        spans = block["spans"]
                    elif "lines" in block:
                        for line in block["lines"]:
                            if "spans" in line:
                                spans.extend(line["spans"])
                    
                    # 如果没有spans，使用原始方法
                    if not spans:
                        return original_process_text_block(self, paragraph, block)
                    
                    # 清除段落现有内容以精确控制格式
                    if paragraph.runs:
                        paragraph.clear()
                    
                    # 处理每个span，精确应用字体样式
                    for span in spans:
                        # 获取文本内容
                        text = span.get("text", "").replace("\u0000", "")
                        if not text.strip():
                            continue
                            
                        # 创建新的run并应用文本
                        run = paragraph.add_run(text)
                        
                        # 应用字体名称
                        if "font" in span and span["font"]:
                            font_name = self._map_font(span["font"])
                            run.font.name = font_name
                        
                        # 应用字体大小
                        if "size" in span and span["size"]:
                            from docx.shared import Pt
                            # 使用更精确的字体大小转换比例
                            size_pt = span["size"] * 0.95  # 调整系数，使Word中显示的字体大小更接近PDF
                            run.font.size = Pt(size_pt)
                        
                        # 应用字体粗细
                        if "flags" in span or "font" in span:
                            font_name = span.get("font", "").lower()
                            flags = span.get("flags", 0)
                            
                            # 检测粗体 - 使用多种检测方法
                            is_bold = False
                            if "bold" in font_name or "heavy" in font_name or "black" in font_name:
                                is_bold = True
                            if flags & 0x20000:  # PDF粗体标志位
                                is_bold = True
                            if "weight" in span and span["weight"] >= 600:  # 字体权重
                                is_bold = True
                                
                            run.bold = is_bold
                            
                            # 检测斜体 - 使用多种检测方法
                            is_italic = False
                            if "italic" in font_name or "oblique" in font_name:
                                is_italic = True
                            if flags & 0x8:  # PDF斜体标志位
                                is_italic = True
                                
                            run.italic = is_italic
                            
                            # 检测下划线
                            if "flags_extra" in span:
                                flags_extra = span.get("flags_extra", 0)
                                if flags_extra & 0x1:  # 下划线标志
                                    run.underline = True
                                if flags_extra & 0x2:  # 删除线标志
                                    run.strike = True
                        
                        # 应用字体颜色
                        if "color" in span and span["color"]:
                            try:
                                from docx.shared import RGBColor
                                color = span["color"]
                                
                                # 处理不同格式的颜色值
                                if isinstance(color, (list, tuple)):
                                    if len(color) >= 3:
                                        # 处理RGB颜色
                                        r, g, b = color[:3]
                                        # 确保颜色值在0-255范围内
                                        if isinstance(r, float) and 0 <= r <= 1:
                                            r = int(r * 255)
                                        if isinstance(g, float) and 0 <= g <= 1:
                                            g = int(g * 255)
                                        if isinstance(b, float) and 0 <= b <= 1:
                                            b = int(b * 255)
                                        run.font.color.rgb = RGBColor(r, g, b)
                                    elif len(color) == 1:  # 灰度值
                                        gray = int(color[0] * 255) if isinstance(color[0], float) else color[0]
                                        run.font.color.rgb = RGBColor(gray, gray, gray)
                                elif isinstance(color, (int, float)):  # 单一灰度值
                                    gray = int(color * 255) if isinstance(color, float) else color
                                    run.font.color.rgb = RGBColor(gray, gray, gray)
                            except Exception as color_err:
                                print(f"应用颜色时出错: {color_err}")
                        
                        # 应用文本间距
                        if "char_spacing" in span and span["char_spacing"]:
                            try:
                                spacing = span["char_spacing"]
                                if spacing != 0:
                                    from docx.oxml.shared import OxmlElement, qn
                                    run_prop = run._element.get_or_add_rPr()
                                    spacing_elm = OxmlElement('w:spacing')
                                    spacing_elm.set(qn('w:val'), str(int(spacing * 20)))
                                    run_prop.append(spacing_elm)
                            except Exception as spacing_err:
                                print(f"应用字符间距时出错: {spacing_err}")
                    
                    return paragraph
                except Exception as e:
                    print(f"增强文本块处理时出错: {e}")
                    traceback.print_exc()
                    # 出错时回退到原始方法
                    return original_process_text_block(self, paragraph, block)
            
            # 保存原始方法并替换为增强版本
            converter._original_process_text_block = original_process_text_block
            converter._process_text_block = types.MethodType(enhanced_process_text_block, converter)
            print("已应用增强文本块处理")
    except Exception as e:
        print(f"增强文本块处理方法失败: {e}")
        traceback.print_exc()

def enhance_paragraph_format_detection(converter):
    """增强段落格式检测，更精确保留段落对齐方式"""
    try:
        # 检查并增强段落格式检测方法
        if hasattr(converter, '_detect_paragraph_format'):
            original_detect_paragraph_format = converter._detect_paragraph_format
            
            def enhanced_detect_paragraph_format(self, block, page_width):
                """增强的段落格式检测方法，更精确识别对齐方式"""
                try:
                    # 检查是否有明确的对齐标记
                    if "alignment" in block:
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        alignment = block["alignment"].lower()
                        if alignment == "center":
                            return WD_ALIGN_PARAGRAPH.CENTER, 0
                        elif alignment == "right":
                            return WD_ALIGN_PARAGRAPH.RIGHT, 0
                        elif alignment == "justify":
                            return WD_ALIGN_PARAGRAPH.JUSTIFY, 0
                        elif alignment == "left":
                            return WD_ALIGN_PARAGRAPH.LEFT, self._detect_left_indent(block)
                    
                    # 没有明确的对齐标记，通过文本位置分析
                    lines = block.get("lines", [])
                    if not lines:
                        return original_detect_paragraph_format(block, page_width)
                    
                    # 提取行的位置信息
                    line_positions = []
                    for line in lines:
                        bbox = line.get("bbox", [0, 0, 0, 0])
                        if len(bbox) >= 4:
                            line_positions.append({
                                "left": bbox[0],
                                "right": bbox[2],
                                "width": bbox[2] - bbox[0],
                                "center": (bbox[0] + bbox[2]) / 2
                            })
                    
                    if not line_positions:
                        return original_detect_paragraph_format(block, page_width)
                    
                    # 计算页面中心位置
                    page_center = page_width / 2
                    
                    # 分析段落对齐方式
                    # 1. 检查是否所有行都接近中心对齐
                    all_center_aligned = True
                    center_tolerance = page_width * 0.05  # 5%的容差
                    
                    for pos in line_positions:
                        line_center = pos["center"]
                        if abs(line_center - page_center) > center_tolerance:
                            all_center_aligned = False
                            break
                    
                    if all_center_aligned:
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        return WD_ALIGN_PARAGRAPH.CENTER, 0
                    
                    # 2. 检查是否所有行都接近右对齐
                    all_right_aligned = True
                    right_tolerance = page_width * 0.05  # 5%的容差
                    
                    for pos in line_positions:
                        line_right = pos["right"]
                        if abs(line_right - page_width) > right_tolerance:
                            all_right_aligned = False
                            break
                    
                    if all_right_aligned:
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        return WD_ALIGN_PARAGRAPH.RIGHT, 0
                    
                    # 3. 检查是否为两端对齐
                    if len(line_positions) > 1:
                        # 获取除最后一行外的所有行宽度
                        other_widths = [pos["width"] for pos in line_positions[:-1]]
                        if other_widths:
                            avg_other_width = sum(other_widths) / len(other_widths)
                            last_line_width = line_positions[-1]["width"]
                            
                            # 如果最后一行明显短于其他行，且其他行宽度接近相等，则可能是两端对齐
                            if last_line_width < avg_other_width * 0.85 and all(abs(w - avg_other_width) < avg_other_width * 0.1 for w in other_widths):
                                from docx.enum.text import WD_ALIGN_PARAGRAPH
                                left_indent = self._detect_left_indent(block)
                                return WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent
                    
                    # 默认回退到原始方法
                    return original_detect_paragraph_format(block, page_width)
                    
                except Exception as e:
                    print(f"增强段落格式检测时出错: {e}")
                    traceback.print_exc()
                    return original_detect_paragraph_format(block, page_width)
            
            # 添加左缩进检测方法
            def detect_left_indent(self, block):
                """检测段落左缩进"""
                try:
                    lines = block.get("lines", [])
                    if not lines:
                        return 0
                    
                    # 获取所有行的左边界
                    left_positions = []
                    for line in lines:
                        bbox = line.get("bbox", [0, 0, 0, 0])
                        if len(bbox) >= 4:
                            left_positions.append(bbox[0])
                    
                    if not left_positions:
                        return 0
                    
                    # 计算左缩进 (将页面坐标转换为Word中的英寸)
                    # 假设页面宽度为8.5英寸 (标准US Letter页面)
                    min_left = min(left_positions)
                    indent_inches = min_left / 72  # 假设PDF单位为点(1/72英寸)
                    
                    # 限制最大缩进
                    max_indent = 2.0  # 最大2英寸的缩进
                    indent_inches = min(indent_inches, max_indent)
                    
                    from docx.shared import Inches
                    return Inches(indent_inches)
                except Exception:
                    return 0
            
            # 保存原始方法并替换为增强版本
            converter._original_detect_paragraph_format = original_detect_paragraph_format
            converter._detect_paragraph_format = types.MethodType(enhanced_detect_paragraph_format, converter)
            converter._detect_left_indent = types.MethodType(detect_left_indent, converter)
            print("已应用增强段落格式检测")
    except Exception as e:
        print(f"增强段落格式检测方法失败: {e}")
        traceback.print_exc()

def enhance_font_mapping(converter):
    """增强字体映射，更准确识别和保留字体"""
    try:
        # 检查并增强字体映射方法
        if hasattr(converter, '_map_font'):
            original_map_font = converter._map_font
            
            def enhanced_map_font(self, pdf_font_name):
                """增强的字体映射方法，更准确识别字体"""
                if not pdf_font_name:
                    return "Arial"
                
                # 尝试导入更高级的字体处理模块
                try:
                    from enhanced_font_handler import map_font
                    return map_font(pdf_font_name, quality="exact")
                except ImportError:
                    pass
                
                # 如果没有高级字体处理模块，使用增强的内置映射
                pdf_font_lower = pdf_font_name.lower().strip()
                
                # 创建扩展的字体映射表
                extended_font_map = {
                    # 基本西文字体
                    "times": "Times New Roman",
                    "times-roman": "Times New Roman",
                    "timesnewroman": "Times New Roman",
                    "times new roman": "Times New Roman",
                    "arial": "Arial",
                    "helvetica": "Arial",
                    "helvetica neue": "Arial",
                    "arial unicode ms": "Arial Unicode MS",
                    "courier": "Courier New",
                    "courier new": "Courier New",
                    "courier-bold": "Courier New",
                    "verdana": "Verdana",
                    "calibri": "Calibri",
                    "calibri light": "Calibri Light",
                    "tahoma": "Tahoma",
                    "georgia": "Georgia",
                    "garamond": "Garamond",
                    "bookman": "Bookman Old Style",
                    "palatino": "Palatino Linotype",
                    "palatino-roman": "Palatino Linotype",
                    "century": "Century Schoolbook",
                    "century schoolbook": "Century Schoolbook",
                    "cambria": "Cambria",
                    "candara": "Candara",
                    "consolas": "Consolas",
                    "constantia": "Constantia",
                    "corbel": "Corbel",
                    "franklin": "Franklin Gothic",
                    "franklin gothic": "Franklin Gothic",
                    "gill": "Gill Sans",
                    "gill sans": "Gill Sans",
                    "lucida": "Lucida Sans",
                    "lucida sans": "Lucida Sans",
                    "segoe ui": "Segoe UI",
                    "segoe": "Segoe UI",
                    "trebuchet": "Trebuchet MS",
                    "trebuchet ms": "Trebuchet MS",
                    
                    # 中文字体
                    "simsun": "SimSun",
                    "songti": "SimSun",
                    "song": "SimSun",
                    "宋体": "SimSun",
                    "simhei": "SimHei",
                    "heiti": "SimHei",
                    "黑体": "SimHei",
                    "microsoft yahei": "Microsoft YaHei",
                    "yahei": "Microsoft YaHei",
                    "微软雅黑": "Microsoft YaHei",
                    "fangsong": "FangSong",
                    "仿宋": "FangSong",
                    "kaiti": "KaiTi",
                    "楷体": "KaiTi",
                    "nsimsun": "NSimSun",
                    "新宋体": "NSimSun",
                    "dfkai": "DFKai-SB",
                    "标楷体": "DFKai-SB",
                    
                    # 日文字体
                    "ms gothic": "MS Gothic",
                    "ms mincho": "MS Mincho",
                    "meiryo": "Meiryo",
                    "yu gothic": "Yu Gothic",
                    "yu mincho": "Yu Mincho",
                    
                    # 韩文字体
                    "malgun gothic": "Malgun Gothic",
                    "gulim": "Gulim",
                    "batang": "Batang",
                    "dotum": "Dotum",
                    "gungsuh": "Gungsuh",
                    
                    # 符号字体
                    "symbol": "Symbol",
                    "wingdings": "Wingdings",
                    "webdings": "Webdings",
                    "zapfdingbats": "Wingdings",
                    "dingbats": "Wingdings",
                }
                
                # 1. 尝试直接匹配
                if pdf_font_lower in extended_font_map:
                    return extended_font_map[pdf_font_lower]
                
                # 2. 尝试部分匹配
                for key, value in extended_font_map.items():
                    if key in pdf_font_lower:
                        return value
                
                # 3. 尝试检测字体族
                if any(x in pdf_font_lower for x in ["sans", "helvetica", "arial"]):
                    return "Arial"
                elif any(x in pdf_font_lower for x in ["serif", "times", "roman"]):
                    return "Times New Roman"
                elif any(x in pdf_font_lower for x in ["mono", "courier", "typewriter"]):
                    return "Courier New"
                elif any(x in pdf_font_lower for x in ["gothic", "gothic", "黑"]):
                    # 检查是否包含中文或日韩字符来确定应该用哪个Gothic字体
                    if any(x in pdf_font_lower for x in ["微软", "yahei", "msyh"]):
                        return "Microsoft YaHei"
                    elif any(x in pdf_font_lower for x in ["ms", "mincho", "明"]):
                        return "MS Gothic"
                    elif "malgun" in pdf_font_lower:
                        return "Malgun Gothic"
                    else:
                        return "Arial"  # 默认西文sans-serif字体
                
                # 4. 回退到原始方法
                try:
                    return original_map_font(self, pdf_font_name)
                except Exception:
                    return "Arial"  # 最终默认字体
            
            # 保存原始方法并替换为增强版本
            converter._original_map_font = original_map_font
            converter._map_font = types.MethodType(enhanced_map_font, converter)
            print("已应用增强字体映射")
    except Exception as e:
        print(f"增强字体映射方法失败: {e}")
        traceback.print_exc()

def enhance_text_position_detection(converter):
    """增强文本位置检测，更精确保留文本布局"""
    try:
        # 添加文本定位方法
        def detect_text_positioning(self, block):
            """检测并保留文本的精确位置信息"""
            position_data = {
                "left": 0,
                "top": 0,
                "width": 0,
                "height": 0,
                "line_spacing": 1.15,  # 默认行间距
                "char_spacing": 0      # 默认字符间距
            }
            
            try:
                # 从block中提取bbox
                if "bbox" in block:
                    bbox = block["bbox"]
                    if len(bbox) >= 4:
                        position_data["left"] = bbox[0]
                        position_data["top"] = bbox[1]
                        position_data["width"] = bbox[2] - bbox[0]
                        position_data["height"] = bbox[3] - bbox[1]
                
                # 分析行间距
                lines = block.get("lines", [])
                if len(lines) > 1:
                    line_tops = []
                    for line in lines:
                        if "bbox" in line and len(line["bbox"]) >= 4:
                            line_tops.append(line["bbox"][1])
                    
                    if len(line_tops) > 1:
                        # 计算平均行间距
                        line_gaps = []
                        for i in range(1, len(line_tops)):
                            line_gaps.append(line_tops[i] - line_tops[i-1])
                        
                        if line_gaps:
                            avg_line_gap = sum(line_gaps) / len(line_gaps)
                            # 估算行间距
                            line_height = position_data["height"] / len(lines)
                            if line_height > 0:
                                line_spacing = avg_line_gap / line_height
                                # 限制在合理范围内
                                if 0.8 <= line_spacing <= 2.5:
                                    position_data["line_spacing"] = line_spacing
                
                # 分析字符间距
                spans = []
                for line in lines:
                    if "spans" in line:
                        spans.extend(line["spans"])
                
                char_spacing_values = []
                for span in spans:
                    if "char_spacing" in span:
                        char_spacing_values.append(span["char_spacing"])
                
                if char_spacing_values:
                    avg_char_spacing = sum(char_spacing_values) / len(char_spacing_values)
                    if abs(avg_char_spacing) > 0.01:  # 忽略非常小的值
                        position_data["char_spacing"] = avg_char_spacing
                
                return position_data
            except Exception as e:
                print(f"检测文本位置时出错: {e}")
                return position_data
        
        # 添加应用文本位置方法
        def apply_text_positioning(self, paragraph, position_data):
            """应用精确的文本位置信息到Word段落"""
            try:
                if not position_data:
                    return
                
                # 应用行间距
                if "line_spacing" in position_data and position_data["line_spacing"] > 0:
                    line_spacing = position_data["line_spacing"]
                    # 转换为Word行间距格式
                    paragraph.paragraph_format.line_spacing = line_spacing
                
                # 应用段落缩进
                if "left" in position_data and position_data["left"] > 0:
                    from docx.shared import Pt
                    # 转换点单位为Word单位
                    left_indent = position_data["left"]
                    paragraph.paragraph_format.left_indent = Pt(left_indent)
                
                # 应用段前间距
                if "top" in position_data and position_data["top"] > 0:
                    from docx.shared import Pt
                    # 只有在段落不是页面第一段时才应用上边距
                    if position_data.get("is_first_paragraph", False) == False:
                        paragraph.paragraph_format.space_before = Pt(position_data["top"] * 0.5)  # 乘以因子调整合适的间距
                
                # 应用段后间距
                if "bottom_margin" in position_data and position_data["bottom_margin"] > 0:
                    from docx.shared import Pt
                    paragraph.paragraph_format.space_after = Pt(position_data["bottom_margin"] * 0.5)  # 乘以因子调整合适的间距
            except Exception as e:
                print(f"应用文本位置时出错: {e}")
        
        # 增强处理段落方法
        if hasattr(converter, '_process_paragraph'):
            original_process_paragraph = converter._process_paragraph
            
            def enhanced_process_paragraph(self, doc, block, page_num=None):
                """增强的段落处理方法，保留精确的文本位置"""
                try:
                    # 检测文本位置信息
                    position_data = self.detect_text_positioning(block)
                    
                    # 调用原始方法创建段落
                    paragraph = original_process_paragraph(self, doc, block, page_num)
                    
                    # 应用位置信息
                    if paragraph:
                        self.apply_text_positioning(paragraph, position_data)
                    
                    return paragraph
                except Exception as e:
                    print(f"增强段落处理时出错: {e}")
                    traceback.print_exc()
                    # 出错时回退到原始方法
                    return original_process_paragraph(self, doc, block, page_num)
            
            # 保存原始方法并替换为增强版本
            converter._original_process_paragraph = original_process_paragraph
            converter._process_paragraph = types.MethodType(enhanced_process_paragraph, converter)
            converter.detect_text_positioning = types.MethodType(detect_text_positioning, converter)
            converter.apply_text_positioning = types.MethodType(apply_text_positioning, converter)
            print("已应用增强文本位置检测")
    except Exception as e:
        print(f"增强文本位置检测方法失败: {e}")
        traceback.print_exc()
