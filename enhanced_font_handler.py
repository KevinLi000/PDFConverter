"""
增强字体处理模块 - 提供增强的字体样式检测和映射
"""

import re

def map_font(pdf_font_name, quality="normal"):
    """
    将PDF字体名称映射到Word字体 - 增强版本
    
    参数:
        pdf_font_name: PDF中的字体名称
        quality: 字体替换质量 ("normal", "high", "exact")
        
    返回:
        Word兼容的字体名称
    """
    # 如果没有字体名称，返回默认字体
    if not pdf_font_name:
        return "Arial"
    
    # 规范化字体名称
    pdf_font_lower = pdf_font_name.lower().strip()
    
    # 根据quality参数选择映射策略
    if quality == "exact":
        return exact_font_mapping(pdf_font_name)
    elif quality == "high":
        return high_quality_font_mapping(pdf_font_name)
    else:  # normal
        return normal_font_mapping(pdf_font_name)

def normal_font_mapping(pdf_font_name):
    """基本字体映射 - 映射常见字体"""
    pdf_font_lower = pdf_font_name.lower().strip()
    
    # 基本字体映射表
    font_map = {
        # 基本字体
        "times": "Times New Roman",
        "times-roman": "Times New Roman",
        "timesnewroman": "Times New Roman",
        "timesnew": "Times New Roman",
        "times new roman": "Times New Roman",
        "roman": "Times New Roman",
        
        # Arial/Helvetica 字体家族
        "arial": "Arial",
        "helvetica": "Arial",
        "helv": "Arial",
        "helveticaneue": "Arial",
        "helvetica neue": "Arial",
        "sans-serif": "Arial",
        "sans serif": "Arial",
        
        # Courier 字体家族
        "courier": "Courier New",
        "couriernew": "Courier New",
        "courier new": "Courier New",
        "cour": "Courier New",
        
        # 其他常见西文字体
        "garamond": "Garamond",
        "book antiqua": "Book Antiqua",
        "bookman": "Bookman Old Style",
        "palatino": "Palatino Linotype",
        "century": "Century Schoolbook",
        "candara": "Candara",
        "consolas": "Consolas",
        "constantia": "Constantia",
        "corbel": "Corbel",
        "franklin": "Franklin Gothic",
        "gill": "Gill Sans",
        "lucida": "Lucida Sans",
        
        # 中文字体
        "simsum": "SimSun",
        "simsun": "SimSun",
        "songti": "SimSun",
        "sim sun": "SimSun",
        "宋体": "SimSun",
        "宋": "SimSun",
        
        "simhei": "SimHei",
        "heiti": "SimHei",
        "sim hei": "SimHei",
        "黑体": "SimHei",
        "黑": "SimHei",
        
        "kaiti": "KaiTi",
        "kai": "KaiTi",
        "楷体": "KaiTi",
        "楷": "KaiTi",
        
        "fangsong": "FangSong",
        "fang song": "FangSong",
        "仿宋": "FangSong",
        
        "msyh": "Microsoft YaHei",
        "microsoft yahei": "Microsoft YaHei",
        "yahei": "Microsoft YaHei",
        "微软雅黑": "Microsoft YaHei",
        "雅黑": "Microsoft YaHei",
        
        "stxihei": "STXihei",
        "华文细黑": "STXihei",
        
        "stkaiti": "STKaiti",
        "华文楷体": "STKaiti",
        
        "stsong": "STSong",
        "华文宋体": "STSong",
        
        # 日语字体
        "ms mincho": "MS Mincho",
        "mincho": "MS Mincho",
        "ms gothic": "MS Gothic",
        "gothic": "MS Gothic",
        "meiryo": "Meiryo",
        
        # 韩语字体
        "batang": "Batang",
        "gulim": "Gulim",
        "malgun gothic": "Malgun Gothic",
        "malgun": "Malgun Gothic",
    }
    
    # 1. 先尝试完全匹配
    if pdf_font_lower in font_map:
        return font_map[pdf_font_lower]
    
    # 2. 部分匹配
    for key, value in font_map.items():
        if key in pdf_font_lower:
            return value
    
    # 3. 智能匹配 - 检查常见字体样式词汇
    is_serif = any(x in pdf_font_lower for x in ["serif", "roman", "times", "ming", "song", "宋"])
    is_sans = any(x in pdf_font_lower for x in ["sans", "arial", "helvetica", "gothic", "hei", "黑"])
    is_mono = any(x in pdf_font_lower for x in ["mono", "courier", "typewriter", "console"])
    
    if is_serif:
        return "Times New Roman"
    elif is_sans:
        return "Arial"
    elif is_mono:
        return "Courier New"
    
    # 默认返回通用字体
    return "Arial"

def high_quality_font_mapping(pdf_font_name):
    """高质量字体映射 - 更精确地映射各种字体变体"""
    pdf_font_lower = pdf_font_name.lower().strip()
    
    # 首先尝试使用常规映射
    result = normal_font_mapping(pdf_font_name)
    if result != "Arial" or pdf_font_lower == "arial":
        return result
    
    # 增强的字体变体检测
    
    # 检测字体粗细
    is_bold = any(x in pdf_font_lower for x in ["bold", "heavy", "black", "strong", "粗", "黑", "dark", "demi"])
    is_light = any(x in pdf_font_lower for x in ["light", "thin", "细", "轻"])
    
    # 检测字体倾斜
    is_italic = any(x in pdf_font_lower for x in ["italic", "oblique", "slant", "斜"])
    
    # 检测宽度变体
    is_condensed = any(x in pdf_font_lower for x in ["condensed", "narrow", "compressed", "紧缩"])
    is_extended = any(x in pdf_font_lower for x in ["extended", "wide", "expanded", "宽"])
    
    # 检测字体类型
    is_serif = any(x in pdf_font_lower for x in ["serif", "roman", "times", "ming", "song", "宋"])
    is_sans = any(x in pdf_font_lower for x in ["sans", "arial", "helvetica", "gothic", "hei", "黑"])
    is_mono = any(x in pdf_font_lower for x in ["mono", "courier", "typewriter", "console"])
    is_script = any(x in pdf_font_lower for x in ["script", "brush", "hand", "cursive", "书法"])
    is_decorative = any(x in pdf_font_lower for x in ["decorative", "ornamental", "display", "fancy", "装饰"])
    
    # 基于特征选择字体
    if is_mono:
        base_font = "Courier New"
    elif is_script:
        base_font = "Script MT" if not is_sans else "Segoe Script"
    elif is_decorative:
        base_font = "Impact" if is_sans else "Old English Text MT"
    elif is_serif:
        if "georgia" in pdf_font_lower:
            base_font = "Georgia"
        elif "cambria" in pdf_font_lower:
            base_font = "Cambria"
        else:
            base_font = "Times New Roman"
    elif is_sans:
        if "verdana" in pdf_font_lower:
            base_font = "Verdana"
        elif "tahoma" in pdf_font_lower:
            base_font = "Tahoma"
        elif "calibri" in pdf_font_lower:
            base_font = "Calibri"
        else:
            base_font = "Arial"
    else:
        # 默认使用Arial
        base_font = "Arial"
    
    return base_font

def exact_font_mapping(pdf_font_name):
    """精确字体映射 - 尝试匹配最精确的字体，包括变体"""
    # 这个函数在实际项目中可能需要一个更完整的字体数据库
    # 这里提供一个简化的实现
    
    pdf_font_lower = pdf_font_name.lower().strip()
    
    # 首先尝试高质量映射
    base_font = high_quality_font_mapping(pdf_font_name)
    
    # 检测更多细节变体
    # 移除已知的字体名称前缀后缀，以检测主要字体族
    cleaned_name = re.sub(r'(regular|std|mt|ms|pro|w\d+|medium|text)', '', pdf_font_lower)
    cleaned_name = cleaned_name.strip(' -_')
    
    # 检测额外的特征
    # 在此处可以添加更多精确匹配的逻辑
    
    # 如果没有更好的匹配，返回高质量映射结果
    return base_font

def detect_font_style(font_info):
    """
    从字体信息中检测字体样式特征
    
    参数:
        font_info: 字体信息字典
        
    返回:
        字体样式信息字典
    """
    style = {
        "bold": False,
        "italic": False,
        "underline": False,
        "strike": False,
        "size": 12,  # 默认大小
        "color": None  # 默认颜色
    }
    
    # 检查字体名称中的样式提示
    if "font" in font_info and font_info["font"]:
        font_name = font_info["font"].lower()
        
        # 检测粗体
        style["bold"] = any(x in font_name for x in ["bold", "heavy", "black", "strong", "粗", "黑", "dark", "demi"])
        
        # 检测斜体
        style["italic"] = any(x in font_name for x in ["italic", "oblique", "slant", "斜"])
    
    # 从字体标志或权重中检测粗体
    if "flags" in font_info:
        flags = font_info["flags"]
        # 一些PDF库使用标志位表示字体样式
        # 通常第1位(0x1)表示固定宽度，第2位(0x2)表示衬线，
        # 第3位(0x4)表示象形文字，第4位(0x8)表示斜体，
        # 第18位(0x20000)表示粗体
        if flags & 0x20000:  # 检查粗体标志
            style["bold"] = True
        if flags & 0x8:  # 检查斜体标志
            style["italic"] = True
    
    if "weight" in font_info:
        # 字体权重通常为100到900，700或以上通常被视为粗体
        weight = font_info["weight"]
        if weight >= 700:
            style["bold"] = True
    
    # 获取字体大小
    if "size" in font_info and font_info["size"]:
        try:
            size = float(font_info["size"])
            if 1 <= size <= 144:  # 合理的字体大小范围
                style["size"] = size
        except (ValueError, TypeError):
            pass
    
    # 获取字体颜色
    if "color" in font_info and font_info["color"]:
        style["color"] = font_info["color"]
    
    # 检测装饰效果
    if "rise" in font_info and font_info["rise"]:
        # 正值表示上标，负值表示下标
        rise = font_info["rise"]
        if rise > 0:
            style["superscript"] = True
        elif rise < 0:
            style["subscript"] = True
    
    # 添加下划线和删除线检测
    if "flags_extra" in font_info:
        flags_extra = font_info["flags_extra"]
        if flags_extra & 0x1:  # 示例：检查下划线标志
            style["underline"] = True
        if flags_extra & 0x2:  # 示例：检查删除线标志
            style["strike"] = True
    
    # 检测特殊的文本装饰标记
    if "font" in font_info and font_info["font"]:
        font_name = font_info["font"].lower()
        if "underline" in font_name or "underlined" in font_name:
            style["underline"] = True
        if "strike" in font_name or "strikethrough" in font_name or "linethrough" in font_name:
            style["strike"] = True
    
    # 检测字距调整
    if "char_spacing" in font_info and font_info["char_spacing"]:
        style["char_spacing"] = font_info["char_spacing"]
    
    # 检测小型大写字母
    if "small_caps" in font_info and font_info["small_caps"]:
        style["small_caps"] = True
    elif "font" in font_info and "smallcaps" in str(font_info["font"]).lower():
        style["small_caps"] = True
    
    return style

def apply_font_style(run, style):
    """
    应用字体样式到Word文档的文本运行
    
    参数:
        run: 文本运行对象
        style: 字体样式信息字典
    """
    # 应用基本样式
    if style.get("bold"):
        run.bold = True
    
    if style.get("italic"):
        run.italic = True
    
    if style.get("underline"):
        run.underline = True
    
    if style.get("strike"):
        run.strike = True
    
    # 应用字体大小
    if "size" in style and style["size"]:
        from docx.shared import Pt
        run.font.size = Pt(style["size"])
    
    # 应用字体颜色
    if "color" in style and style["color"]:
        from docx.shared import RGBColor
        color = style["color"]
        if isinstance(color, tuple) and len(color) == 3:
            r, g, b = color
            run.font.color.rgb = RGBColor(r, g, b)
    
    # 应用上下标
    if style.get("superscript"):
        run.font.superscript = True
    if style.get("subscript"):
        run.font.subscript = True
