#!/usr/bin/env python
"""
PDF字体管理模块
用于处理PDF字体映射和替换
"""

import sys
import os
import traceback

class PDFFontManager:
    """PDF字体管理器，处理字体映射和替换"""
    
    def __init__(self):
        """初始化字体管理器"""
        self.font_substitution_quality = "normal"  # normal, high, exact
        self.font_mapping = {}  # 字体映射表
        self.system_fonts = []  # 系统可用字体
        self.custom_fonts = []  # 自定义字体
        self.force_font_embedding = False  # 是否强制嵌入字体
        self.initialized = False
        
        # 初始化字体映射
        self._initialize_font_mapping()
    
    def _initialize_font_mapping(self):
        """初始化字体映射表"""
        # 基本字体映射
        self.font_mapping = {
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
           
            # 其他常用字体
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
        
        # 标记初始化完成
        self.initialized = True
    
    def set_quality(self, quality="normal"):
        """
        设置字体替换质量
        
        参数:
            quality: 字体替换质量 (normal, high, exact)
        """
        if quality in ["normal", "high", "exact"]:
            self.font_substitution_quality = quality
        else:
            print(f"警告: 不支持的字体替换质量: {quality}，使用默认值'normal'")
            self.font_substitution_quality = "normal"
    
    def add_font_mapping(self, source_font, target_font):
        """
        添加自定义字体映射
        
        参数:
            source_font: 源字体名称
            target_font: 目标字体名称
        """
        if source_font and target_font:
            self.font_mapping[source_font.lower().strip()] = target_font
    
    def map_font(self, pdf_font_name, fallback_font="Arial"):
        """
        将PDF字体名称映射到可用字体
        
        参数:
            pdf_font_name: PDF中的字体名称
            fallback_font: 备用字体
            
        返回:
            映射后的字体名称
        """
        if not pdf_font_name:
            return fallback_font
            
        # 确保映射表已初始化
        if not self.initialized:
            self._initialize_font_mapping()
            
        # 转换为小写便于匹配
        pdf_font_lower = pdf_font_name.lower().strip()
        
        # 1. 检查是否有直接匹配
        if pdf_font_lower in self.font_mapping:
            return self.font_mapping[pdf_font_lower]
        
        # 2. 如果质量设置为高或精确，尝试部分匹配
        if self.font_substitution_quality in ["high", "exact"]:
            for key, value in self.font_mapping.items():
                if key in pdf_font_lower or pdf_font_lower in key:
                    return value
        
        # 3. 如果质量设置为精确，进行智能匹配
        if self.font_substitution_quality == "exact":
            # 检查常见字体样式词汇
            is_serif = any(x in pdf_font_lower for x in ["serif", "roman", "times", "ming", "song", "宋"])
            is_sans = any(x in pdf_font_lower for x in ["sans", "arial", "helvetica", "gothic", "hei", "黑"])
            is_mono = any(x in pdf_font_lower for x in ["mono", "courier", "typewriter", "console"])
            
            if is_serif:
                return "Times New Roman"
            elif is_sans:
                return "Arial"
            elif is_mono:
                return "Courier New"
        
        # 返回默认字体
        return fallback_font
    
    def get_font_style(self, pdf_font_name):
        """
        从字体名称中检测字体样式
        
        参数:
            pdf_font_name: PDF中的字体名称
            
        返回:
            字体样式信息 (is_bold, is_italic)
        """
        if not pdf_font_name:
            return False, False
            
        pdf_font_lower = pdf_font_name.lower().strip()
        
        # 检测是否为粗体
        is_bold = any(x in pdf_font_lower for x in 
                     ["bold", "black", "heavy", "strong", "粗体", "粗", "黑体", "黑"])
        
        # 检测是否为斜体
        is_italic = any(x in pdf_font_lower for x in 
                       ["italic", "oblique", "slant", "斜体", "斜"])
        
        return is_bold, is_italic
    
    def scan_system_fonts(self):
        """扫描系统可用字体"""
        try:
            # 不同操作系统的字体路径
            font_paths = []
            
            if sys.platform.startswith('win'):
                # Windows字体路径
                font_paths.append(os.path.join(os.environ['WINDIR'], 'Fonts'))
            elif sys.platform.startswith('darwin'):
                # macOS字体路径
                font_paths.extend([
                    '/Library/Fonts',
                    '/System/Library/Fonts',
                    os.path.expanduser('~/Library/Fonts')
                ])
            elif sys.platform.startswith('linux'):
                # Linux字体路径
                font_paths.extend([
                    '/usr/share/fonts',
                    '/usr/local/share/fonts',
                    os.path.expanduser('~/.fonts')
                ])
            
            # 扫描字体文件
            system_fonts = []
            for font_path in font_paths:
                if os.path.exists(font_path) and os.path.isdir(font_path):
                    for file in os.listdir(font_path):
                        if file.lower().endswith(('.ttf', '.ttc', '.otf')):
                            font_name = os.path.splitext(file)[0]
                            system_fonts.append(font_name)
            
            self.system_fonts = system_fonts
            return system_fonts
            
        except Exception as e:
            print(f"扫描系统字体时出错: {e}")
            traceback.print_exc()
            return []
