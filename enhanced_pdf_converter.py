#!/usr/bin/env python
"""
增强型PDF转换工具 - 将PDF转换为Word和Excel文件，精确保留原始格式
作者: GitHub Copilot
日期: 2025-05-27
"""

import os
import sys
import io
import re
import argparse
import tempfile
import shutil
import traceback
import numpy as np
import pandas as pd
import cv2
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.shared import OxmlElement, qn
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.table import _Cell, Table
from docx.styles import styles
from docxtpl import DocxTemplate
# 修复tabula导入问题，确保在不同版本的tabula-py中都能正常工作
import tabula
# 确保read_pdf可用 - 这是为了解决'can't import name read_pdf from tabula'错误
# 在一些tabula-py版本中，read_pdf应该从tabula.io导入
try:
    from tabula.io import read_pdf
    tabula.read_pdf = read_pdf
except (ImportError, AttributeError):
    # 如果导入失败，但tabula已经导入成功，那么read_pdf可能已经是tabula的属性
    if not hasattr(tabula, 'read_pdf'):
        print("警告: 无法确定tabula.read_pdf的来源，可能影响表格提取功能")
from PIL import Image, ImageOps, ImageEnhance, ImageCms, ImageCms
import camelot
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, Color, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import BarChart, Reference, LineChart
from openpyxl.worksheet.table import Table as XLTable, TableStyleInfo
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
from pathlib import Path
from collections import Counter, defaultdict

# 尝试导入集成辅助模块
try:
    import converter_integration
    has_integration_helpers = True
except ImportError:
    has_integration_helpers = False

# 尝试导入修复模块
try:
    import pdf_converter_fix
    has_converter_fix = True
except ImportError:
    has_converter_fix = False

# 尝试以不同方式导入PyMuPDF
try:
    import fitz  # PyMuPDF
except ImportError:
    try:
        import PyMuPDF as fitz
    except ImportError:
        print("错误: 无法导入PyMuPDF库，请使用以下命令安装:")
        print("pip install PyMuPDF")
        sys.exit(1)

class EnhancedPDFConverter:
    """增强型PDF转换工具类，精确保留PDF原始格式"""    
    def __init__(self):
        """初始化转换器"""
        self.pdf_path = None
        self.output_dir = None
        self.temp_dir = None        
        self._dpi = 300  # 使用私有变量存储DPI值
        
        # 增强格式保留的配置参数
        self.format_preservation_level = "standard"  # standard, enhanced, maximum
        self.exact_layout_preservation = False  # 精确布局保留
        self.preserve_vector_graphics = True  # 保留矢量图形
        self.detect_tables_accurately = True  # 精确表格检测
        self.smart_color_management = True  # 智能颜色管理
        self.font_substitution_quality = "normal"  # normal, high, exact
        self.image_compression_quality = 95  # 图像压缩质量 (1-100)
        self.force_image_for_tables = True  # 表格区域强制使用图像模式
        self.force_font_embedding = True  # 强制嵌入字体
        self.layout_tolerance = 5  # 布局识别容差值(越小越精确)
        
        # 初始化专用的格式保留管理器
        try:
            # 应用高级表格修复
            self._init_advanced_table_fixes()
        except Exception as e:
            print(f"初始化高级表格修复失败: {e}")

    def _is_complex_page(self, page):
            """检测页面是否包含复杂内容"""
            # 获取页面内容统计
            text = page.get_text()
            blocks = page.get_text("dict")["blocks"]
            
            # 增强图像检测 - 使用多种方法检测图像
            image_blocks = []
            
            # 方法1: 基于块类型检测图像
            basic_image_blocks = [b for b in blocks if b["type"] == 1]
            image_blocks.extend(basic_image_blocks)
            
            # 方法2: 使用get_images方法检测嵌入图像
            try:
                embedded_images = page.get_images()
                if embedded_images and len(embedded_images) > 0:
                    # 将嵌入图像的信息添加到图像块列表
                    for img in embedded_images:
                        # 检查这个图像是否已经在基本图像块中
                        xref = img[0]
                        already_detected = any(b.get("xref") == xref for b in basic_image_blocks)
                        
                        if not already_detected:
                            # 创建一个表示此图像的块
                            image_blocks.append({
                                "type": 1,  # 图像类型
                                "xref": xref,
                                "is_additional_image": True
                            })
            except Exception as e:
                print(f"使用get_images方法检测图像时出错: {e}")
            
            # 检查是否有图像
            has_images = len(image_blocks) > 0
            
            # 检查文本块数量
            text_blocks = [b for b in blocks if b["type"] == 0]
            many_text_blocks = len(text_blocks) > 15
            
            # 检查是否有表格
            # TableFinder对象不支持len()操作，改用其他方法检测表格
            try:
                # 使用表格特征检测是否存在表格
                tables_dict = page.find_tables().extract()
                has_tables = len(tables_dict) > 0
            except:
                # 备用方法：检查页面文本是否包含表格特征
                text_lower = text.lower()
                table_indicators = ['table', '表格', '列表', 'column', 'row', '行', '列']
                table_structure = text.count('|') > 5 or text.count('\t') > 5
                has_tables = any(indicator in text_lower for indicator in table_indicators) or table_structure
            
            # 检查是否有复杂布局
            has_complex_layout = False
            
            # 分析文本块的位置分布
            if len(text_blocks) > 5:
                # 收集所有文本块的x坐标
                x_positions = []
                for block in text_blocks:
                    x_positions.append(block["bbox"][0])  # 左边界
                
                # 如果x坐标分布在多个不同位置，可能是多列布局
                x_pos_counter = Counter(int(x // 20) * 20 for x in x_positions)  # 按20点为间隔分组
                distinct_x_pos = len([k for k, v in x_pos_counter.items() if v > 2])  # 至少出现3次的x位置
                has_complex_layout = distinct_x_pos >= 3
            
            # 增强格式保留模式下更积极地判定为复杂页面
            if hasattr(self, 'format_preservation_level'):
                if self.format_preservation_level == "maximum":
                    # 最大保留模式 - 只要有任何一个复杂因素就判定为复杂
                    return has_images or has_tables or has_complex_layout or many_text_blocks
                elif self.format_preservation_level == "enhanced":
                    # 增强保留模式 - 至少两个复杂因素
                    complexity_factors = sum([has_images, has_tables, has_complex_layout, many_text_blocks])
                # print(f"初始化高级表格修复失败: {e}")
            
            try:
                # 使用模块集成器获取颜色管理器
                import pdf_module_integrator
                self.color_manager = pdf_module_integrator.get_color_manager()
                self._has_color_manager = self.color_manager is not None
            except ImportError as e:
                print(f"无法导入模块集成器: {e}")
                self.color_manager = None
                # self._has_color_manager = Falsetry:
                # 使用相对导入或绝对导入路径
                try:
                    from pdf_font_manager import PDFFontManager
                except ImportError:
                    # 尝试使用绝对导入
                    import sys
                    import os
                    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
                    from pdf_font_manager import PDFFontManager
                self.font_manager = PDFFontManager()
                self._has_font_manager = True
            except ImportError as e:
                print(f"无法导入字体管理器: {e}")
                self._has_font_manager = False
                # 应用修复和增强
            try:
                import pdf_converter_fix
                pdf_converter_fix.apply_enhanced_pdf_converter_fixes(self)
                print("已应用PDF转换器增强修复")
            except ImportError:
                print("修复模块不可用，将尝试加载替代方法")
                
            # 尝试加载表格检测功能 - 优先使用增强型表格检测
            try:
                from enhanced_table_detection import apply_enhanced_table_detection_patch
                apply_enhanced_table_detection_patch(self)
                print("已加载增强型表格检测功能")
            except ImportError:
                # 尝试加载基础表格检测功能
                try:
                    from table_detection_utils import add_table_detection_capability
                    add_table_detection_capability(self)
                    print("已加载基础表格检测功能")
                except ImportError:
                    print("无法导入表格检测工具，可能影响表格识别功能")
            
    @property
    def dpi(self):
        """DPI属性的getter"""
        return self._dpi
        
    @dpi.setter
    def dpi(self, value):
        """DPI属性的setter，确保值为整数"""
        try:
            self._dpi = int(value)  # 确保DPI值为整数
        except (ValueError, TypeError):
            print(f"警告: DPI值必须为整数，获取到: {value}，使用默认值300")
            self._dpi = 300    
    def enhance_format_preservation(self):
        """启用增强的格式保留模式，使用最佳设置确保最高精度的格式保留"""
        self.format_preservation_level = "maximum"
        self.exact_layout_preservation = True
        self.preserve_vector_graphics = True
        self.detect_tables_accurately = True
        self.smart_color_management = True
        self.font_substitution_quality = "exact"
        self.force_image_for_complex_tables = True  # 只对复杂表格使用图像
        self.force_font_embedding = True
        self.layout_tolerance = 1  # 更严格的布局容差以提高精度
        
        # 增加DPI以确保高质量图像渲染 - 确保使用整数
        self.dpi = max(int(self.dpi), 800)  # 提高默认DPI
        
        # 启用智能页面布局分析
        self.enable_smart_layout_analysis = True
        
        # 精确的段落和行间距保留
        self.preserve_exact_spacing = True
        
        # 精确的表格单元格内容对齐
        self.exact_table_cell_alignment = True
        
        # 设置更精确的文本提取选项
        self.text_extraction_mode = "precise"
        self.line_gap_detection = "enhanced"
        self.preserve_text_positioning = True
        
        # 启用高级字体处理
        self.advanced_font_handling = True
        self.font_style_detection = "precise"
        self.character_spacing_preservation = True
        
        # 集成增强色彩和字体管理
        self._initialize_enhanced_managers()
        
        print("已启用最大格式保留模式，将尽可能精确地保留原始PDF格式")
        return self
        
    def _initialize_enhanced_managers(self):
        """初始化增强的格式管理器"""        # 尝试初始化色彩管理器
        try:
            # 使用相对导入或绝对导入路径
            try:
                from pdf_color_manager import PDFColorManager
            except ImportError:
                # 尝试使用绝对导入
                import sys
                import os
                sys.path.append(os.path.dirname(os.path.abspath(__file__)))
                from pdf_color_manager import PDFColorManager
            self.color_manager = PDFColorManager()
            self._has_color_manager = True
            print("已加载增强色彩管理功能")
        except ImportError as e:
            self._has_color_manager = False
            print(f"注意: 未能加载增强色彩管理模块: {e}")
              # 尝试初始化字体管理器
        try:
            # 使用相对导入或绝对导入路径
            try:
                from pdf_font_manager import PDFFontManager
            except ImportError:
                # 尝试使用绝对导入
                import sys
                import os
                sys.path.append(os.path.dirname(os.path.abspath(__file__)))
                from pdf_font_manager import PDFFontManager
            self.font_manager = PDFFontManager()
            self.font_manager.font_substitution_quality = self.font_substitution_quality
            self._has_font_manager = True
            print("已加载增强字体管理功能")
        except ImportError as e:
            self._has_font_manager = False
            print(f"注意: 未能加载增强字体管理模块: {e}")
    
    def set_paths(self, pdf_path, output_dir=None):
        """设置PDF路径和输出目录"""
        self.pdf_path = pdf_path
        
        if output_dir is None:
            # 如果未指定输出目录，使用PDF所在目录
            self.output_dir = os.path.dirname(pdf_path)
        else:
            self.output_dir = output_dir
            
        # 确保输出目录存在
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 创建临时目录
        self.temp_dir = tempfile.mkdtemp()
    
    def cleanup(self):
        """清理临时文件"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def pdf_to_word(self, method="advanced"):
        """
        将PDF转换为Word文档，精确保留原始格式和样式
        
        参数:
            method (str): 转换方法，可选值:
                - "basic": 基本转换，只提取文本
                - "hybrid": 混合模式，同时使用文本提取和图像渲染
                - "advanced": 高级模式，保留最精确的格式(默认)
        """
        if not self.pdf_path:
            raise ValueError("未设置PDF路径")
        
        if method == "basic":
            return self._pdf_to_word_basic()
        elif method == "hybrid":
            return self._pdf_to_word_hybrid()
        else:  # advanced
            return self._pdf_to_word_advanced()
            
    def _pdf_to_word_advanced(self):
        """高级PDF到Word转换，使用页面图像渲染，确保最精确的格式保留"""
        # 创建Word文档
        doc = Document()
        
        try:
            # 打开PDF文件
            pdf_document = fitz.open(self.pdf_path)
            
            # 获取页面数量
            page_count = len(pdf_document)
            
            # 设置页面大小和边距，使用更精确的页面尺寸
            if page_count > 0:
                # 获取第一页的尺寸来设置文档属性
                first_page = pdf_document[0]
                page_width = first_page.rect.width
                page_height = first_page.rect.height
                
                # 判断页面方向
                is_landscape = page_width > page_height
                
                # 设置页面大小和方向
                section = doc.sections[0]
                if is_landscape:
                    section.orientation = WD_ORIENT.LANDSCAPE
                    section.page_width = Cm(29.7)
                    section.page_height = Cm(21)
                else:
                    section.orientation = WD_ORIENT.PORTRAIT
                    section.page_width = Cm(21)
                    section.page_height = Cm(29.7)
                
                # 设置较小的页边距以最大化可用空间
                section.left_margin = Cm(1.0)
                section.right_margin = Cm(1.0)
                section.top_margin = Cm(1.0)
                section.bottom_margin = Cm(1.0)
            else:
                # 如果PDF没有页面，使用默认设置
                section = doc.sections[0]
                section.page_width = Cm(21)
                section.page_height = Cm(29.7)
                section.left_margin = Cm(1.5)
                section.right_margin = Cm(1.5)
                section.top_margin = Cm(1.5)
                section.bottom_margin = Cm(1.5)
            
            # 处理每一页
            for page_num in range(page_count):
                page = pdf_document[page_num]
                
                # 渲染页面为图像并添加到文档
                self._render_page_as_image(doc, page)
                
                # 添加分页符（除了最后一页）
                if page_num < page_count - 1:
                    doc.add_page_break()
            
            # 生成输出文件路径
            pdf_filename = os.path.basename(self.pdf_path)
            output_filename = os.path.splitext(pdf_filename)[0] + ".docx"
            output_path = os.path.join(self.output_dir, output_filename)
            
            # 保存Word文档
            doc.save(output_path)
            
            print(f"成功将PDF转换为Word(高级模式): {output_path}")
            return output_path
            
        except Exception as e:
            print(f"PDF转Word失败: {str(e)}")
            raise
        finally:
            self.cleanup()
    
    def _render_page_as_image(self, doc, page):
        """将整个页面渲染为高质量图像并添加到文档，确保完全精确保留原始格式"""
        # 根据格式保留级别设置缩放比例
        if hasattr(self, 'format_preservation_level') and self.format_preservation_level == "maximum":
            zoom = 12  # 超高质量
        elif hasattr(self, 'format_preservation_level') and self.format_preservation_level == "enhanced":
            zoom = 10  # 增强质量
        else:
            zoom = 8  # 标准质量
        
        # 计算基于DPI的缩放比例 - 确保使用整数DPI值
        dpi_zoom = int(self.dpi) / 72.0  # 72 DPI是PDF的标准分辨率
        zoom = max(zoom, dpi_zoom)  # 使用较大的缩放比例
        
        # 使用高级渲染选项
        render_options = {
            "alpha": True,  # 包含透明度
            "colorspace": fitz.csRGB,  # RGB色彩空间
        }
        
        # 渲染为高分辨率图像
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, **render_options)
        
        # 保存为临时图像
        img_path = os.path.join(self.temp_dir, f"page_hq_{page.number}.png")
        pix.save(img_path, output="png")  # 保存为PNG格式以保持最高质量
        
        # 使用PIL进行图像优化
        try:
            with Image.open(img_path) as img:
                # 应用图像增强
                if hasattr(self, 'smart_color_management') and self.smart_color_management:
                    # 增强对比度
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(1.08)  # 轻微增强对比度
                    
                    # 增强清晰度
                    enhancer = ImageEnhance.Sharpness(img)
                    img = enhancer.enhance(1.2)
                
                # 保存优化后的图像
                img.save(img_path, format='PNG', optimize=True, 
                       quality=self.image_compression_quality if hasattr(self, 'image_compression_quality') else 95)
        except Exception as e:
            print(f"图像优化失败，使用原始渲染: {e}")
        
        # 获取页面尺寸并精确计算Word文档中的图像尺寸
        width_inches = page.rect.width / 72.0  # 转换为英寸
        
        # 确保图像尺寸适应Word页面，同时保留原始宽高比
        max_width_inches = 6.5  # 默认最大宽度
        try:
            # 获取当前部分的可用宽度
            section_width = doc.sections[0].page_width.inches
            margins = doc.sections[0].left_margin.inches + doc.sections[0].right_margin.inches
            max_width_inches = section_width - margins - 0.1  # 减去0.1英寸的安全边距
        except:
            pass
        
        # 添加图像到文档
        try:
            # 使用精确宽度的图片添加方式
            img_width = min(width_inches, max_width_inches)
            doc.add_picture(img_path, width=Inches(img_width))
            
            # 为图像添加精确的对齐方式
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        except Exception as img_err:
            print(f"添加图像时出错: {img_err}，尝试备用方法")
            # 备用方法：添加无格式的图像
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run()
            r.add_picture(img_path, width=Inches(min(width_inches, max_width_inches)))    
    def _map_font(self, pdf_font_name):
        """将PDF字体名称映射到Word字体 - 增强版本"""
        try:
            # 尝试导入增强字体处理模块
            from enhanced_font_handler import map_font
            return map_font(pdf_font_name, quality=self.font_substitution_quality if hasattr(self, 'font_substitution_quality') else "normal")
        except ImportError:
            # 如果增强模块不可用，使用内置方法
            return self._map_font_internal(pdf_font_name)
    
    def _map_font_internal(self, pdf_font_name):
        """内置的字体映射方法"""
        # 如果没有字体名称，返回默认字体
        if not pdf_font_name:
            return "Arial"
            
        # 转换为小写便于匹配
        pdf_font_lower = pdf_font_name.lower().strip()
        
        # 详细的字体映射表
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
        
        # 检查是否有直接匹配
        pdf_font_lower = pdf_font_name.lower().strip()
        
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
    
    def _detect_paragraph_format(self, block, page_width):
        """
        检测文本块的段落格式（对齐方式和缩进）
        
        参数:
            block: 文本块
            page_width: 页面宽度
            
        返回:
            tuple: (alignment, left_indent) - 对齐方式和左缩进值
        """
        # 获取块的边界框
        bbox = block["bbox"]
        left = bbox[0]
        right = bbox[2]
        width = right - left
        
        # 获取块中所有的行
        lines = block.get("lines", [])
        if not lines:
            return WD_ALIGN_PARAGRAPH.LEFT, 0
        
        # 收集所有行的左右边界
        line_lefts = []
        line_rights = []
        line_widths = []
        
        for line in lines:
            line_bbox = line["bbox"]
            line_left = line_bbox[0]
            line_right = line_bbox[2]
            line_width = line_right - line_left
            
            line_lefts.append(line_left)
            line_rights.append(line_right)
            line_widths.append(line_width)
        
        # 计算平均值
        avg_left = sum(line_lefts) / len(line_lefts)
        avg_right = sum(line_rights) / len(line_rights)
        avg_width = sum(line_widths) / len(line_widths)
        
        # 页面中央位置
        page_center = page_width / 2
        
        # 计算文本块中心点
        block_center = (avg_left + avg_right) / 2
        
        # 检测左缩进
        left_indent = 0
        if avg_left > 20:  # 如果左边距大于20点，认为有缩进
            left_indent = avg_left
        
        # 检查是否为居中对齐
        center_tolerance = page_width * 0.1  # 10%的页面宽度作为容差
        if abs(block_center - page_center) < center_tolerance:
            # 额外检查：如果文本宽度很小（相对于页面），更可能是居中的
            if avg_width < page_width * 0.7:  # 文本宽度小于页面宽度的70%
                return WD_ALIGN_PARAGRAPH.CENTER, 0
        
        # 检查是否为右对齐
        right_margin = page_width - avg_right
        if right_margin < 50 and avg_left > 100:  # 右边距小，左边距大
            return WD_ALIGN_PARAGRAPH.RIGHT, 0
        
        # 检查是否为两端对齐（判断标准：多行文本，且最后一行明显短于其他行）
        if len(lines) > 1:
            # 获取除最后一行外的所有行宽度
            other_line_widths = line_widths[:-1]
            if other_line_widths:
                avg_other_width = sum(other_line_widths) / len(other_line_widths)
                last_line_width = line_widths[-1]
                
                # 如果最后一行明显短于其他行（小于80%），可能是两端对齐
                if last_line_width < avg_other_width * 0.8 and avg_width > page_width * 0.7:
                    return WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent
        
        # 检查是否有特殊的段落样式标记
        try:
            spans = []
            for line in lines:
                for span in line.get("spans", []):
                    spans.append(span)
            
            # 检查是否包含居中的标题特征（粗体、大字体等）
            if spans:
                first_span = spans[0]
                font_size = first_span.get("size", 0)
                font_flags = first_span.get("flags", 0)
                
                # 粗体 (0x1)、大字体 (> 12)、居中位置，很可能是标题
                if (font_flags & 0x1) and font_size > 12 and abs(block_center - page_center) < center_tolerance:
                    return WD_ALIGN_PARAGRAPH.CENTER, 0
        except Exception as e:
            print(f"分析段落样式时出错: {e}")
        
        # 默认为左对齐，返回检测到的左缩进
        return WD_ALIGN_PARAGRAPH.LEFT, left_indent
    def _is_new_paragraph_by_indent(self, block, current_paragraph):
        """
        通过缩进判断是否需要新段落
        
        参数:
            block: 当前文本块
            current_paragraph: 当前段落对象
        
        返回:
            bool: 如果需要新段落返回True，否则返回False
        """
        # 如果没有当前段落，总是创建新段落
        if current_paragraph is None:
            return True
            
        # 获取当前块的边界框
        bbox = block["bbox"]
        left_x = bbox[0]  # 左边界X坐标
        
        # 默认缩进差异阈值（以点为单位）
        indent_threshold = 5
        
        # 获取当前块内所有行的左边界位置
        try:
            lines = block.get("lines", [])
            if not lines:
                return True  # 如果没有行信息，创建新段落
                
            # 获取第一行的左边界位置
            first_line_x = lines[0]["bbox"][0]
            
            # 获取当前段落的最后一个运行对象的内容和属性
            try:
                paragraph_content = current_paragraph.text
                
                # 如果段落为空，则认为需要新段落
                if not paragraph_content.strip():
                    return True
                
                # 检查段落末尾是否有结束标志
                if paragraph_content.rstrip().endswith(('.', '!', '?', ':', ';', '。', '！', '？', '：', '；')):
                    return True
                
                # 检查段落是否以不完整的单词结束（可能是断行）
                words = paragraph_content.split()
                if words and len(words[-1]) <= 2:  # 短词可能是断词
                    return False  # 可能是同一段落的延续
                
                # 检查缩进差异
                # 获取段落的左缩进（如果有的话）
                try:
                    paragraph_indent = current_paragraph.paragraph_format.left_indent
                    if paragraph_indent:
                        paragraph_indent = paragraph_indent.pt  # 转换为点
                    else:
                        paragraph_indent = 0
                        
                    # 计算缩进差异
                    indent_diff = abs(first_line_x - paragraph_indent)
                    
                    # 如果缩进差异大于阈值，创建新段落
                    if indent_diff > indent_threshold:
                        return True
                except:
                    # 如果无法获取段落缩进，使用简单的启发式方法
                    pass
                    
                # 检查段落最后一行是否已满（如果不满，可能是段落中间断行）
                # 这是一个启发式规则：如果段落的最后一行很短，可能是新段落的开始
                if len(paragraph_content.rstrip()) < 50:  # 假设少于50个字符表示行未满
                    # 检查当前块是否有明显的缩进
                    if first_line_x > 20:  # 有明显缩进
                        return True
            except:
                # 如果无法分析段落内容，使用简单规则
                pass
                
            # 分析当前块的文本风格
            spans = []
            for line in lines:
                for span in line.get("spans", []):
                    spans.append(span)
                    
            # 如果块中有特殊的格式标记（如粗体、斜体等），可能是新段落的开始
            if spans and any(span.get("flags", 0) > 0 for span in spans):
                return True
                
            # 检查块的第一个字符是否为首字母大写（英文）或中文段落开始的标志
            first_chars = []
            for span in spans:
                if span.get("text"):
                    first_chars.append(span["text"][0])
                    break
                    
            if first_chars and any(c.isupper() for c in first_chars):
                # 首字母大写可能表示新段落（英文）
                return True
                
        except Exception as e:
            print(f"分析段落缩进时出错: {e}")
            return True  # 出错时保守处理，创建新段落
        
        # 默认情况下，如果无法明确判断，则不创建新段落
        return False
    
    def _detect_paragraph_format(self, block, page_width):
        """
        检测文本块的段落格式（对齐方式和缩进）
        
        参数:
            block: 文本块
            page_width: 页面宽度
            
        返回:
            tuple: (alignment, left_indent) - 对齐方式和左缩进值
        """
        # 获取块的边界框
        bbox = block["bbox"]
        left = bbox[0]
        right = bbox[2]
        width = right - left
        
        # 获取块中所有的行
        lines = block.get("lines", [])
        if not lines:
            return WD_ALIGN_PARAGRAPH.LEFT, 0
        
        # 收集所有行的左右边界
        line_lefts = []
        line_rights = []
        line_widths = []
        
        for line in lines:
            line_bbox = line["bbox"]
            line_left = line_bbox[0]
            line_right = line_bbox[2]
            line_width = line_right - line_left
            
            line_lefts.append(line_left)
            line_rights.append(line_right)
            line_widths.append(line_width)
        
        # 计算平均值
        avg_left = sum(line_lefts) / len(line_lefts)
        avg_right = sum(line_rights) / len(line_rights)
        avg_width = sum(line_widths) / len(line_widths)
        
        # 页面中央位置
        page_center = page_width / 2
        
        # 计算文本块中心点
        block_center = (avg_left + avg_right) / 2
        
        # 检测左缩进
        left_indent = 0
        if avg_left > 20:  # 如果左边距大于20点，认为有缩进
            left_indent = avg_left
        
        # 检查是否为居中对齐
        center_tolerance = page_width * 0.1  # 10%的页面宽度作为容差
        if abs(block_center - page_center) < center_tolerance:
            # 额外检查：如果文本宽度很小（相对于页面），更可能是居中的
            if avg_width < page_width * 0.7:  # 文本宽度小于页面宽度的70%
                return WD_ALIGN_PARAGRAPH.CENTER, 0
        
        # 检查是否为右对齐
        right_margin = page_width - avg_right
        if right_margin < 50 and avg_left > 100:  # 右边距小，左边距大
            return WD_ALIGN_PARAGRAPH.RIGHT, 0
        
        # 检查是否为两端对齐（判断标准：多行文本，且最后一行明显短于其他行）
        if len(lines) > 1:
            # 获取除最后一行外的所有行宽度
            other_line_widths = line_widths[:-1]
            if other_line_widths:
                avg_other_width = sum(other_line_widths) / len(other_line_widths)
                last_line_width = line_widths[-1]
                
                # 如果最后一行明显短于其他行（小于80%），可能是两端对齐
                if last_line_width < avg_other_width * 0.8 and avg_width > page_width * 0.7:
                    return WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent
        
        # 检查是否有特殊的段落样式标记
        try:
            spans = []
            for line in lines:
                for span in line.get("spans", []):
                    spans.append(span)
            
            # 检查是否包含居中的标题特征（粗体、大字体等）
            if spans:
                first_span = spans[0]
                font_size = first_span.get("size", 0)
                font_flags = first_span.get("flags", 0)
                
                # 粗体 (0x1)、大字体 (> 12)、居中位置，很可能是标题
                if (font_flags & 0x1) and font_size > 12 and abs(block_center - page_center) < center_tolerance:
                    return WD_ALIGN_PARAGRAPH.CENTER, 0
        except Exception as e:
            print(f"分析段落样式时出错: {e}")
        
        # 默认为左对齐，返回检测到的左缩进
        return WD_ALIGN_PARAGRAPH.LEFT, left_indent

    def _init_advanced_table_fixes(self):
        """初始化高级表格修复功能"""
        try:
            # 尝试导入高级表格修复模块
            import advanced_table_fixes
            
            # 应用高级表格修复
            self = advanced_table_fixes.apply_advanced_table_fixes(self)
            
            print("已初始化高级表格修复")
            
        except ImportError:
            print("高级表格修复模块不可用，将跳过此功能")
        except Exception as e:
            print(f"应用高级表格修复时出错: {e}")
            traceback.print_exc()    
    def _process_text_block_enhanced(self, paragraph, block):
        """
        处理文本块，增强版 - 尽量保留段落格式
        参数:
            paragraph: Word文档中的段落对象
            block: PDF文本块
        """
        lines = block.get("lines", [])
        if not lines:
            return

        # 检测对齐和缩进
        page_width = block.get("page_width", 595)  # 取默认A4宽度
        align, left_indent = self._detect_paragraph_format(block, page_width)
        paragraph.alignment = align
        # Clamp left_indent to a safe range (0-100 points)
        if left_indent > 0:
            left_indent = min(max(left_indent, 0), 100)
            paragraph.paragraph_format.left_indent = Pt(left_indent * 0.35)  # 点转磅

        # 检测标题
        is_heading = False
        if "heading" in str(block.get("type", "")).lower() or any(
            span.get("size", 0) > 14 for line in lines for span in line.get("spans", [])
        ):
            is_heading = True
            paragraph.style = "Heading 1"

        # 检测列表
        first_text = lines[0]["spans"][0]["text"].strip() if lines and lines[0].get("spans") else ""
        if first_text.startswith(("-", "•", "·")):
            paragraph.style = "List Bullet"
        elif first_text[:2].isdigit() and first_text[2:3] in (".", "、"):
            paragraph.style = "List Number"

        # 处理文本内容，保留换行符
        prev_left = None
        prev_y_bottom = None
        for idx, line in enumerate(lines):
            line_text = "".join(span.get("text", "") for span in line.get("spans", []))
            current_y_top = line["bbox"][1]
            
            if idx == 0:
                # 第一行直接添加
                paragraph.add_run(line_text)
                prev_left = line["bbox"][0]
                prev_y_bottom = line["bbox"][3]  # 底部y坐标
            else:
                # 判断是否新段落（如缩进/空行/行距大等）
                is_new_para = False
                line_spacing = current_y_top - prev_y_bottom if prev_y_bottom else 0
                
                # 检查明显的段落变化
                if abs(line["bbox"][0] - prev_left) > 10:  # 缩进变化
                    is_new_para = True
                elif line_spacing > 15:  # 行间距明显大于普通行间距
                    is_new_para = True
                elif not line_text.strip():  # 空行
                    is_new_para = True
                
                if is_new_para:
                    # 创建新段落
                    paragraph = paragraph._parent.add_paragraph()
                    paragraph.alignment = align
                    if left_indent > 0:
                        paragraph.paragraph_format.left_indent = Pt(left_indent * 0.35)
                    paragraph.add_run(line_text)
                else:
                    # 同一段落内的换行
                    # 添加换行符并继续在同一段落
                    last_run = paragraph.runs[-1] if paragraph.runs else None
                    if last_run:
                        # 检查上一个run的文本是否以换行结束
                        if not last_run.text.endswith('\n'):
                            last_run.add_break()  # 添加换行符
                    paragraph.add_run(line_text)
                
                prev_left = line["bbox"][0]
                prev_y_bottom = line["bbox"][3]

    def _pdf_to_word_basic(self):
        """基本的PDF到Word转换，增强版本，更精确保留原始格式和样式"""
        # 创建Word文档
        doc = Document()
        
        try:
            # 打开PDF文件
            pdf_document = fitz.open(self.pdf_path)
            
            # 获取第一页的尺寸用于设置文档默认属性
            if len(pdf_document) > 0:
                first_page = pdf_document[0]
                page_width = first_page.rect.width
                page_height = first_page.rect.height
                is_landscape = page_width > page_height
            else:
                # 默认A4尺寸
                page_width = 595  # A4宽度点数
                page_height = 842  # A4高度点数
                is_landscape = False
            
            # 设置页面大小和边距
            section = doc.sections[0]
            if is_landscape:
                section.orientation = WD_ORIENT.LANDSCAPE
                section.page_width = Cm(29.7)
                section.page_height = Cm(21)
            else:
                section.orientation = WD_ORIENT.PORTRAIT
                section.page_width = Cm(21)
                section.page_height = Cm(29.7)
            
            # 设置更精确的页边距以匹配PDF原始边距
            margin = min(20, max(10, page_width * 0.05))  # 根据页面宽度自适应边距
            section.left_margin = Cm(margin / 28.35)    # 将点转换为厘米
            section.right_margin = Cm(margin / 28.35)
            section.top_margin = Cm(margin / 28.35)
            section.bottom_margin = Cm(margin / 28.35)            # 预先检测文档中的表格
            tables_by_page = {}
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                try:
                    # 加载并使用增强的表格检测功能
                    if not hasattr(self, 'detect_tables'):
                        try:
                            # 尝试导入并添加表格检测功能
                            from table_detection_utils import add_table_detection_capability
                            add_table_detection_capability(self)
                            print("已加载增强表格检测功能")
                        except ImportError:
                            print("无法导入表格检测工具，尝试使用内置功能")
                    
                    # 使用表格检测方法
                    if hasattr(self, 'detect_tables'):
                        # 使用增强的detect_tables方法
                        tables = self.detect_tables(page)
                        if tables:
                            if hasattr(tables, 'tables'):  # PyMuPDF的find_tables结果
                                tables_by_page[page_num] = tables.tables
                            else:  # 自定义表格检测结果
                                tables_by_page[page_num] = tables
                    else:
                        # 尝试使用PyMuPDF的内置方法（可能不存在）
                        try:
                            tables = page.find_tables()
                            if tables and len(tables.tables) > 0:
                                tables_by_page[page_num] = tables.tables
                        except AttributeError:
                            # 如果find_tables不可用，尝试使用备用方法
                            if hasattr(self, '_extract_tables') or hasattr(self, '_extract_tables_fallback'):
                                extract_func = getattr(self, '_extract_tables', None) or getattr(self, '_extract_tables_fallback')
                                tables = extract_func(pdf_document, page_num)
                                if tables:
                                    tables_by_page[page_num] = tables
                except Exception as table_err:
                    print(f"表格检测警告 (页 {page_num+1}): {table_err}")
            
            # 检测是否有多列布局的页面
            multi_column_pages = self._detect_multi_column_pages(pdf_document)
            
            # 遍历PDF页面
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                
                # 分析页面布局
                page_dict = page.get_text("dict", sort=True)  # 使用sort=True确保按阅读顺序
                blocks = page_dict["blocks"]
                
                # 预处理块，标记表格区域
                blocks = self._mark_table_regions(blocks, tables_by_page.get(page_num, []))
                
                # 检测页面方向，如果当前页与文档默认设置不同，需要添加新的节并设置方向
                if page_num > 0:
                    # 获取当前页尺寸
                    curr_width = page.rect.width
                    curr_height = page.rect.height
                    curr_is_landscape = curr_width > curr_height
                    
                    # 如果当前页方向与前一页不同，添加新节
                    if curr_is_landscape != is_landscape:
                        doc.add_section(WD_SECTION.NEW_PAGE)
                        section = doc.sections[-1]
                        if curr_is_landscape:
                            section.orientation = WD_ORIENT.LANDSCAPE
                            section.page_width = Cm(29.7)
                            section.page_height = Cm(21)
                        else:
                            section.orientation = WD_ORIENT.PORTRAIT
                            section.page_width = Cm(21)
                            section.page_height = Cm(29.7)
                        
                        # 保持相同的边距设置
                        section.left_margin = Cm(margin / 28.35)
                        section.right_margin = Cm(margin / 28.35)
                        section.top_margin = Cm(margin / 28.35)
                        section.bottom_margin = Cm(margin / 28.35)
                        
                        # 更新当前页面方向状态
                        is_landscape = curr_is_landscape
                
                # 获取当前页面内容的行列结构
                column_positions = self._detect_columns(blocks)
                line_positions = self._detect_lines(blocks)
                
                # 检查当前页是否是多列布局
                is_multi_column = page_num in multi_column_pages
                columns_count = multi_column_pages.get(page_num, 1)
                
                # 如果是多列布局，为当前页面添加节并设置多列
                if is_multi_column and columns_count > 1:
                    # 如果不是第一页，需要先添加新节
                    if page_num > 0:
                        doc.add_section(WD_SECTION.NEW_PAGE)
                        section = doc.sections[-1]
                        
                        # 保持当前页面方向
                        if is_landscape:
                            section.orientation = WD_ORIENT.LANDSCAPE
                            section.page_width = Cm(29.7)
                            section.page_height = Cm(21)
                        else:
                            section.orientation = WD_ORIENT.PORTRAIT
                            section.page_width = Cm(21)
                            section.page_height = Cm(29.7)
                        
                        # 保持边距设置
                        section.left_margin = Cm(margin / 28.35)
                        section.right_margin = Cm(margin / 28.35)
                        section.top_margin = Cm(margin / 28.35)
                        section.bottom_margin = Cm(margin / 28.35)
                    else:
                        section = doc.sections[0]
                    
                    # 设置多列
                    sectPr = section._sectPr
                    cols = OxmlElement('w:cols')
                    cols.set(qn('w:num'), str(columns_count))
                    # 添加列间距设置
                    cols.set(qn('w:space'), "425")  # 约0.3英寸的列间距
                    sectPr.append(cols)
                
                # 处理页面内容的方式取决于是否为多列布局
                if is_multi_column:
                    # 对于多列布局，按列分别处理内容
                    self._process_multi_column_page(doc, page, pdf_document, blocks, column_positions)
                else:
                    # 对于单列布局，按常规方式处理
                    # 按y0坐标排序块，以保持垂直阅读顺序
                    blocks.sort(key=lambda b: (b["bbox"][1], b["bbox"][0]))
                    current_y = -1
                    current_paragraph = None
                    for block in blocks:
                        if block.get("is_table", False):
                            self._process_table_block(doc, block, page, pdf_document)
                            current_paragraph = None
                            current_y = -1
                            continue
                        if block["type"] == 1:
                            self._process_image_block_enhanced(doc, pdf_document, page, block)
                            current_paragraph = None
                            current_y = -1
                            continue
                        if block["type"] == 0:
                            block_y = block["bbox"][1]
                            new_paragraph_needed = (current_y == -1 or 
                                                   (abs(block_y - current_y) > 12) or  
                                                   self._is_new_paragraph_by_indent(block, current_paragraph))
                            if new_paragraph_needed:
                                current_paragraph = doc.add_paragraph()
                                current_y = block_y
                                try:
                                    format_result = self._detect_paragraph_format(block, page.rect.width)
                                    if isinstance(format_result, tuple) and len(format_result) == 2:
                                        alignment, left_indent = format_result
                                    else:
                                        alignment = WD_ALIGN_PARAGRAPH.LEFT
                                        left_indent = 0
                                    current_paragraph.alignment = alignment
                                    # Clamp left_indent to a safe range (0-100 points)
                                    if left_indent > 0:
                                        left_indent = min(max(left_indent, 0), 100)
                                        current_paragraph.paragraph_format.left_indent = Cm(left_indent / 28.35)
                                except Exception as e:
                                    print(f"设置段落格式时出错: {e}")
                                    current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

                            # 处理文本内容 - 使用增强的文本处理函数
                            self._process_text_block_enhanced(current_paragraph, block)
                
                # 如果不是最后一页，添加分页符
                if page_num < len(pdf_document) - 1:
                    doc.add_page_break()
            
            # 生成输出文件路径
            pdf_filename = os.path.basename(self.pdf_path)
            output_filename = os.path.splitext(pdf_filename)[0] + ".docx"
            output_path = os.path.join(self.output_dir, output_filename)
            
            # 保存Word文档
            doc.save(output_path)
            
            print(f"成功将PDF转换为Word: {output_path}")            
            return output_path
            
        except Exception as e:
            print(f"PDF转Word失败: {str(e)}")
            raise
        finally:
            self.cleanup()
            
    def _add_table_as_image(self, doc, page, bbox):
        """
        将表格区域作为图像添加到Word文档，确保表格可见
        
        参数:
            doc: Word文档对象
            page: PDF页面
            bbox: 表格边界框
        """
        try:
            import fitz
            from docx.shared import Inches
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            
            # 使用高分辨率渲染表格区域
            clip_rect = fitz.Rect(bbox)
            zoom = 3.0  # 高分辨率
            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, clip=clip_rect, alpha=False)
            
            # 保存为临时文件
            import os
            image_path = os.path.join(self.temp_dir, f"table_image_{page.number}_{hash(str(bbox))}.png")
            pix.save(image_path)
            
            # 添加图像到文档
            if os.path.exists(image_path):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # 计算表格宽度并设置图像宽度
                table_width = (bbox[2] - bbox[0]) / 72.0  # 转换为英寸（假设72 DPI）
                max_width = 6.0  # 最大宽度（英寸）
                
                # 添加图像
                p.add_run().add_picture(image_path, width=Inches(min(max_width, table_width)))
                print(f"成功将表格作为图像添加: {image_path}")
                
                # 添加一个空段落作为间距
                doc.add_paragraph()
        except Exception as e:
            print(f"将表格作为图像添加时出错: {e}")
            import traceback
            traceback.print_exc()    
    def _validate_and_fix_table_data(self, table_data, merged_cells=None):
        """
        验证表格数据并修复常见问题
        
        参数:
            table_data: 表格数据二维列表
            merged_cells: 合并单元格信息列表，每项为 (start_row, start_col, end_row, end_col)
        
        返回:
            修复后的表格数据和合并单元格信息
        """
        if not table_data:
            return [], []
            
        # 确保表格数据有效
        if not isinstance(table_data, list):
            print("警告: 表格数据不是列表类型")
            return [], []
            
        # 确保表格至少有一行
        if len(table_data) == 0:
            return [], []
            
        # 检查行一致性
        col_count = 0
        for row in table_data:
            if isinstance(row, list):
                col_count = max(col_count, len(row))
                
        if col_count == 0:
            print("警告: 表格没有有效列")
            return [], []
            
        # 初始化修复后的表格数据
        fixed_table_data = []
        
        # 确保所有行具有相同的列数
        for row_idx, row in enumerate(table_data):
            if not isinstance(row, list):
                # 如果行不是列表，创建一个空行
                fixed_row = [""] * col_count
            else:
                # 确保行长度一致
                fixed_row = list(row)
                if len(fixed_row) < col_count:
                    # 填充缺失的单元格
                    fixed_row.extend([""] * (col_count - len(fixed_row)))
                elif len(fixed_row) > col_count:
                    # 截断过长的行
                    fixed_row = fixed_row[:col_count]
            
            # 处理单元格内容
            for i in range(len(fixed_row)):
                cell_content = fixed_row[i]
                
                # 将None替换为空字符串
                if cell_content is None:
                    fixed_row[i] = ""
                
                # 处理非字符串类型
                if not isinstance(cell_content, str):
                    try:
                        fixed_row[i] = str(cell_content)
                    except:
                        fixed_row[i] = ""
                
                # 处理多行文本 - 确保保留换行符
                if isinstance(fixed_row[i], str):
                    # 替换连续空格为单个空格，但保留换行符
                    fixed_row[i] = re.sub(r' {2,}', ' ', fixed_row[i])
                    # 删除行首行尾空白，但保留内部格式
                    fixed_row[i] = fixed_row[i].strip()
            
            fixed_table_data.append(fixed_row)
          # 验证合并单元格信息
        if merged_cells is None:
            merged_cells = []
        
        fixed_merged_cells = []
        for merge_info in merged_cells:
            if (isinstance(merge_info, (list, tuple)) and 
                len(merge_info) == 4 and 
                all(isinstance(idx, int) for idx in merge_info)):
                start_row, start_col, end_row, end_col = merge_info
                
                # 确保索引在有效范围内
                if (0 <= start_row <= end_row < len(fixed_table_data) and
                    0 <= start_col <= end_col < col_count):
                    fixed_merged_cells.append((start_row, start_col, end_row, end_col))
        
                    return fixed_table_data, fixed_merged_cells
            for col_idx in range(len(fixed_row)):
                cell_value = fixed_row[col_idx]
                
                # 转换None为空字符串
                if cell_value is None:
                    fixed_row[col_idx] = ""
                    continue
                    
                # 尝试将单元格内容转换为字符串
                try:
                    # 如果是数字，保留原值便于后续格式化
                    if isinstance(cell_value, (int, float)):
                        fixed_row[col_idx] = cell_value
                    else:
                        # 转换为字符串并去除前后空白
                        fixed_row[col_idx] = str(cell_value).strip()
                except Exception as e:
                    print(f"转换单元格内容时出错 ({row_idx}, {col_idx}): {e}")
                    fixed_row[col_idx] = ""
            
            fixed_table_data.append(fixed_row)
        
        # 验证并修复合并单元格信息
        fixed_merged_cells = []
        if merged_cells:
            row_count = len(fixed_table_data)
            
            for merge_info in merged_cells:
                if len(merge_info) != 4:
                    print(f"警告: 无效的合并单元格信息: {merge_info}")
                    continue
                    
                start_row, start_col, end_row, end_col = merge_info
                
                # 确保索引在有效范围内
                start_row = max(0, min(start_row, row_count - 1))
                end_row = max(start_row, min(end_row, row_count - 1))
                start_col = max(0, min(start_col, col_count - 1))
                end_col = max(start_col, min(end_col, col_count - 1))
                
                # 添加有效的合并单元格信息
                fixed_merged_cells.append((start_row, start_col, end_row, end_col))
        
        # 处理空表格的特殊情况
        if len(fixed_table_data) == 0:
            # 创建一个最小的有效表格 (1x1)
            fixed_table_data = [["无数据"]]
            print("警告: 创建了默认的空表格")
        
        # 检测并修复无效字符
        for row_idx, row in enumerate(fixed_table_data):
            for col_idx, cell_value in enumerate(row):
                if isinstance(cell_value, str):
                    # 替换控制字符和其他无效字符
                    clean_value = ''.join(c if (c.isprintable() or c in ['\n', '\t']) else ' ' for c in cell_value)
                    
                    # 处理过长的单元格内容
                    if len(clean_value) > 32767:  # Word单元格文本长度限制
                        clean_value = clean_value[:32764] + "..."
                        print(f"警告: 单元格 ({row_idx}, {col_idx}) 内容过长，已截断")
                    
                    fixed_table_data[row_idx][col_idx] = clean_value
        
        return fixed_table_data, fixed_merged_cells
    def _process_table_block(self, doc, block, page, pdf_document):
        """
        处理表格块并添加到Word文档 - 完整保留表格样式和结构
        """
        try:
            # 优先使用增强型表格检测
            try:
                from enhanced_table_detection import extract_table_data
                table_data, merged_cells = extract_table_data(block, page)
            except Exception:                
                table_data = block.get("table_data", [])
                merged_cells = block.get("merged_cells", [])
            fixed_table_data, fixed_merged_cells = self._validate_and_fix_table_data(table_data, merged_cells)
            rows = len(fixed_table_data)
            cols = len(fixed_table_data[0]) if rows > 0 else 0
            if rows == 0 or cols == 0:
                self._add_table_as_image(doc, page, block["bbox"])
                return

            # 检测并应用表格样式
            try:
                from enhanced_table_style import detect_table_style, apply_cell_style, apply_table_style
                table_style_info = detect_table_style(block, page)
                use_enhanced_style = True
            except Exception:
                table_style_info = {}
                use_enhanced_style = False

            # 创建表格
            word_table = doc.add_table(rows=rows, cols=cols)
            word_table.style = table_style_info.get("table_style", "Table Grid")

            # 应用增强边框和内边距
            try:
                from improved_table_borders import apply_enhanced_borders
                border_width = table_style_info.get("border_width", 8)
                border_color = table_style_info.get("border_color", "000000")
                if isinstance(border_color, (tuple, list)):
                    border_color = "%02x%02x%02x" % tuple(border_color[:3])
                apply_enhanced_borders(word_table, border_width, border_color)
            except Exception as border_err:
                print(f"表格边框增强失败: {border_err}")            # 合并单元格
            for merge_info in fixed_merged_cells:
                start_row, start_col, end_row, end_col = merge_info                
                if (0 <= start_row < rows and 0 <= start_col < cols and 0 <= end_row < rows and 0 <= end_col < cols):
                    try:
                        cell = word_table.cell(start_row, start_col)
                        # 先进行垂直合并
                        if end_row > start_row:
                            try:
                                cell.merge(word_table.cell(end_row, start_col))
                            except Exception as e:
                                print(f"垂直合并单元格时出错: {e}")
                        
                        # 再进行水平合并
                        if end_col > start_col:
                            try:
                                # 如果前面的垂直合并成功，则使用合并后的单元格
                                # 否则使用原始单元格
                                merged_cell = word_table.cell(start_row, start_col)
                                merged_cell.merge(word_table.cell(start_row, end_col))
                            except Exception as e:
                                print(f"水平合并单元格时出错: {e}")
                    except Exception as e:
                        print(f"处理合并单元格时出错: {e} - 跳过此合并操作")# 填充内容并应用样式
            for i, row in enumerate(fixed_table_data):
                for j, cell_content in enumerate(row):
                    cell = word_table.cell(i, j)
                    cell.text = ''
                    # 多行文本处理
                    if isinstance(cell_content, str) and '\n' in cell_content:
                        lines = cell_content.split('\n')
                        if lines[0]:
                            first_paragraph = cell.paragraphs[0]
                            first_paragraph.add_run(lines[0])
                        for line in lines[1:]:
                            if line.strip():
                                new_paragraph = cell.add_paragraph()
                                new_paragraph.add_run(line)
                                new_paragraph.space_before = 0
                                new_paragraph.space_after = 0
                    else:
                        if cell_content is not None and str(cell_content).strip():
                            cell.paragraphs[0].add_run(str(cell_content))
                    # 应用单元格样式
                    if use_enhanced_style:
                        try:
                            apply_cell_style(cell, table_style_info, i, j)
                        except Exception as style_err:
                            print(f"应用单元格样式失败: {style_err}")
            # 列宽设置
            if table_style_info.get("col_widths"):
                for idx, col_width in enumerate(table_style_info["col_widths"]):
                    if idx < len(word_table.columns):
                        try:
                            if col_width and int(col_width) > 0:
                                word_table.columns[idx].width = int(col_width)
                        except Exception:
                            pass
            # 斑马纹
            if table_style_info.get("zebra_striping"):
                from docx.shared import RGBColor
                alt_color = table_style_info.get("alternate_row_color", (240,240,240))
                for i, row in enumerate(word_table.rows):
                    if i % 2 == 1:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                para.runs[0].font.highlight_color = None
                            shading_elm = cell._element.get_or_add_tcPr()
                            from docx.oxml import parse_xml
                            from docx.oxml.ns import nsdecls
                            color_hex = "%02x%02x%02x" % tuple(alt_color[:3])
                            shading_elm.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'))
            # 垂直居中
            for row in word_table.rows:
                for cell in row.cells:
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            # 表格整体样式
            if use_enhanced_style:
                try:
                    apply_table_style(word_table, table_style_info)
                except Exception as e:
                    print(f"应用表格整体样式时出错: {e}")
            doc.add_paragraph().space_after = Pt(6)
        except Exception as e:
            print(f"处理表格时出错: {e}，使用图像备用方案")
            import traceback
            traceback.print_exc()
            self._add_table_as_image(doc, page, block["bbox"])
    def _process_image_block_enhanced(self, doc, pdf_document, page, block):
        """
        处理图像块，修复版 - 确保图像正确显示
        
        参数:
            doc: Word文档对象
            pdf_document: PDF文档对象
            page: 页面对象
            block: 图像块
        """
        try:
            xref = block.get("xref", 0)
            bbox = block["bbox"]
            page_width = page.rect.width
            image_left = bbox[0]
            image_right = bbox[2]
            image_width = image_right - image_left
            image_height = bbox[3] - bbox[1]

            # 提取图像数据 - 增强图像识别流程
            img = None
            img_path = None
            extraction_methods = []
            
            # 方法1: 通过xref提取嵌入图片
            if xref:
                try:
                    pix = fitz.Pixmap(pdf_document, xref)
                    if pix.n > 4:  # 处理带alpha通道的图像
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    img_path = os.path.join(self.temp_dir, f"image_{page.number}_{xref}.png")
                    pix.save(img_path)
                    if os.path.exists(img_path):
                        extraction_methods.append(("xref", img_path))
                except Exception as e:
                    print(f"通过xref提取图片失败: {e}")
            
            # 方法2: 通过bbox裁剪页面区域，使用更高分辨率
            try:
                clip_rect = fitz.Rect(bbox)
                zoom = 4.0  # 提高分辨率 (原为2.0)
                matrix = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=matrix, clip=clip_rect, alpha=False)
                img_path = os.path.join(self.temp_dir, f"image_{page.number}_{hash(str(bbox))}_high_res.png")
                pix.save(img_path)
                if os.path.exists(img_path):
                    extraction_methods.append(("bbox_high_res", img_path))
            except Exception as e:
                print(f"通过高分辨率bbox裁剪图片失败: {e}")
            
            # 方法3: 尝试使用更大的边界框
            try:
                # 扩大边界框以捕获可能被错误裁剪的图像
                expanded_bbox = [
                    max(0, bbox[0] - 5),
                    max(0, bbox[1] - 5),
                    min(page.rect.width, bbox[2] + 5),
                    min(page.rect.height, bbox[3] + 5)
                ]
                clip_rect = fitz.Rect(expanded_bbox)
                zoom = 3.0
                matrix = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=matrix, clip=clip_rect, alpha=False)
                img_path = os.path.join(self.temp_dir, f"image_{page.number}_{hash(str(expanded_bbox))}_expanded.png")
                pix.save(img_path)
                if os.path.exists(img_path):
                    extraction_methods.append(("expanded_bbox", img_path))
            except Exception as e:
                print(f"通过扩展边界框裁剪图片失败: {e}")
            
            # 选择最佳图像
            if not extraction_methods:
                print("未能提取到图片，跳过该图像块")
                return
            
            # 选择最大的图像文件（通常质量更好）
            selected_img_path = max(extraction_methods, 
                                   key=lambda x: os.path.getsize(x[1]) if os.path.exists(x[1]) else 0)[1]
            
            # 重新计算图像尺寸以确保正确的宽高比
            try:
                section_width = doc.sections[0].page_width.inches
                margins = doc.sections[0].left_margin.inches + doc.sections[0].right_margin.inches
                max_width_inches = section_width - margins - 0.1
            except:
                max_width_inches = 6.0
            img_width = (image_width / 72.0) if image_width > 0 else max_width_inches
            img_width = min(img_width, max_width_inches)

            # 插入图片
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(img_path, width=Inches(img_width))
            doc.add_paragraph()  # 增加间距
        except Exception as img_err:
            print(f"插入图片时出错: {img_err}")

    def _mark_table_regions(self, blocks, tables):
        """
        标记属于表格区域的块 - 兼容不同表格对象格式
        
        参数:
            blocks: 页面的内容块
            tables: 在页面中检测到的表格列表
        
        返回:
            更新后的 块列表，带有表格标记
        """
        if not tables:
            return blocks
            
        # 复制块列表以避免修改原始数据
        marked_blocks = []
        
        # 将表格转换为块并标记
        for table in tables:
            table_rect = None
            
            # 1. 尝试获取表格矩形区域
            try:
                # 方法1: 直接访问rect属性 (PyMuPDF 1.18.0+)
                if hasattr(table, 'rect'):
                    table_rect = table.rect
                # 方法2: 直接访问bbox属性
                elif hasattr(table, 'bbox'):
                    table_rect = fitz.Rect(table.bbox)
                # 方法3: 字典对象的bbox属性
                elif isinstance(table, dict) and "bbox" in table:
                    table_rect = fitz.Rect(table["bbox"])
                # 方法4: 从单元格计算表格范围
                else:
                    cells = None
                    
                    # 获取单元格列表
                    if hasattr(table, 'cells') and table.cells:
                        cells = table.cells
                    elif isinstance(table, dict) and "cells" in table and table["cells"]:
                        cells = table["cells"]
                    elif hasattr(table, 'tables') and table.tables and len(table.tables) > 0:
                        first_table = table.tables[0]
                        if hasattr(first_table, 'cells') and first_table.cells:
                            cells = first_table.cells
                    
                    if cells and len(cells) > 0:
                        # 从单元格计算表格范围
                        bboxes = []
                        
                        for cell in cells:
                            cell_bbox = None
                            
                            if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                                cell_bbox = cell["bbox"]
                            elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                                cell_bbox = cell[:4]
                            elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                                cell_bbox = cell.bbox
                                
                            if cell_bbox:
                                bboxes.append(cell_bbox)
                        
                        if bboxes:
                            min_x = min(bbox[0] for bbox in bboxes)
                            min_y = min(bbox[1] for bbox in bboxes)
                            max_x = max(bbox[2] for bbox in bboxes)
                            max_y = max(bbox[3] for bbox in bboxes)
                            table_rect = fitz.Rect(min_x, min_y, max_x, max_y)
                
                # 如果无法获取表格区域，跳过此表格
                if not table_rect:
                    print("警告: 无法获取表格区域，跳过此表格")
                    continue
            except Exception as e:
                print(f"警告: 处理表格边界时出错: {e}")
                continue
                
            # 2. 提取并修正表格数据
            try:
                table_data = []
                merged_cells = []                  # 方法1: 使用extract方法
                if hasattr(table, 'extract'):
                    merged_cells_info = self._detect_merged_cells(table)
                    extracted_data = table.extract()
                    table_data, merged_cells = self._validate_and_fix_table_data(extracted_data, merged_cells_info)
                # 方法2: 字典中已包含table_data
                elif isinstance(table, dict) and "table_data" in table:
                    table_data = table["table_data"]
                    merged_cells = table.get("merged_cells", [])
                # 方法3: 从单元格构建表格
                else:
                    table_data, merged_cells = self._build_table_from_cells(table)
            except Exception as e:
                print(f"警告: 提取表格数据时出错: {e}")
                import traceback
                traceback.print_exc()
                table_data = []
                merged_cells = []
            
            # 跳过空表格
            if not table_data:
                continue
                
            # 创建表格块
            table_block = {
                "type": 100,  # 自定义类型表示表格
                "bbox": [table_rect.x0, table_rect.y0, table_rect.x1, table_rect.y1],
                "is_table": True,
                "table_data": table_data,
                "merged_cells": merged_cells,
                "rows": len(table_data),
                "cols": len(table_data[0]) if table_data and table_data[0] else 0
            }
            
            marked_blocks.append(table_block)
        
        # 添加非表格区域的块
        for block in blocks:
            block_rect = fitz.Rect(block["bbox"])
            
            # 检查此块是否与任何表格重叠
            is_in_table = False
            for table_block in [b for b in marked_blocks if b.get("is_table", False)]:
                table_rect = fitz.Rect(table_block["bbox"])
                # 检查重叠
                if block_rect.intersects(table_rect) and block_rect.get_area() / block_rect.get_area() > 0.5:
                    is_in_table = True
                    break
                    
            # 如果不在表格中，添加到最终块列表
            if not is_in_table:
                marked_blocks.append(block)
        
        # 按垂直位置排序
        marked_blocks.sort(key=lambda b: b["bbox"][1])
        return marked_blocks

    def _build_table_from_cells(self, table):
        """
        从单元格数据构建表格结构
        
        参数:
            table: 表格对象
            
        返回:
            构建的表格数据和合并单元格信息
        """
        # 检查是否是字典对象
        if isinstance(table, dict):
            # 如果是字典对象，尝试从table_data获取数据
            if "table_data" in table and isinstance(table["table_data"], list):
                return table["table_data"], table.get("merged_cells", [])
            # 检查字典是否包含cells属性
            elif "cells" in table and table["cells"]:
                cells = table["cells"]
            else:
                return [], []
        # 检查是否有cells属性
        elif hasattr(table, 'cells') and table.cells:
            cells = table.cells
        # 如果表格有tables属性 (PyMuPDF 1.18.0+)
        elif hasattr(table, 'tables') and table.tables:
            # 尝试获取第一个表格
            if len(table.tables) > 0:
                first_table = table.tables[0]
                if hasattr(first_table, 'cells') and first_table.cells:
                    cells = first_table.cells
                else:
                    return [], []
            else:
                return [], []
        else:
            return [], []
            
        try:
            # 识别行和列的位置
            row_positions = set()
            col_positions = set()
            
            for cell in cells:
                # 收集所有行和列的起始位置
                if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                    # 处理字典形式的单元格
                    bbox = cell["bbox"]
                    row_positions.add(bbox[1])  # 上边界
                    row_positions.add(bbox[3])  # 下边界
                    col_positions.add(bbox[0])  # 左边界
                    col_positions.add(bbox[2])  # 右边界
                elif isinstance(cell, (list, tuple)) and len(cell) >= 4:  # 确保单元格有足够的坐标信息
                    row_positions.add(cell[1])  # 上边界
                    row_positions.add(cell[3])  # 下边界
                    col_positions.add(cell[0])  # 左边界
                    col_positions.add(cell[2])  # 右边界
                elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                    row_positions.add(cell.bbox[1])  # 上边界
                    row_positions.add(cell.bbox[3])  # 下边界
                    col_positions.add(cell.bbox[0])  # 左边界
                    col_positions.add(cell.bbox[2])  # 右边界
                    
            # 排序位置
            row_positions = sorted(row_positions)
            col_positions = sorted(col_positions)
            
            # 创建空表格
            rows_count = len(row_positions) - 1
            cols_count = len(col_positions) - 1
            
            if rows_count <= 0 or cols_count <= 0:
                return [], []
                
            # 初始化表格和占位标记矩阵
            table_data = [["" for _ in range(cols_count)] for _ in range(rows_count)]
            occupied = [[False for _ in range(cols_count)] for _ in range(rows_count)]
            merged_cells = []  # 存储合并单元格信息: (行开始, 列开始, 行结束, 列结束)
            
            # 为每个单元格创建映射，以便查找其在表格中的位置
            cell_position_map = {}
            
            # 首先识别所有单元格的位置
            for cell in cells:
                # 获取单元格坐标
                cell_bbox = None
                cell_text = ""
                
                if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                    # 处理字典形式的单元格
                    cell_bbox = cell["bbox"]
                    cell_text = cell.get("text", "")
                elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                    cell_bbox = cell[:4]
                    if len(cell) > 4 and isinstance(cell[4], str):
                        cell_text = cell[4]
                elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                    cell_bbox = cell.bbox
                    if hasattr(cell, 'text'):
                        cell_text = cell.text
                
                if not cell_bbox:
                    continue
                    
                left, top, right, bottom = cell_bbox[0], cell_bbox[1], cell_bbox[2], cell_bbox[3]
                
                # 找出单元格在表格网格中的位置
                row_start = row_positions.index(top) if top in row_positions else -1
                row_end = row_positions.index(bottom) if bottom in row_positions else -1
                col_start = col_positions.index(left) if left in col_positions else -1
                col_end = col_positions.index(right) if right in col_positions else -1
                
                # 跳过无效位置
                if row_start < 0 or row_end <= row_start or col_start < 0 or col_end <= col_start:
                    continue
                
                # 存储单元格位置信息
                cell_key = (left, top, right, bottom)
                cell_position_map[cell_key] = (row_start, col_start, row_end, col_end, cell_text)
            
            # 然后填充表格内容并识别合并单元格
            for cell_key, position_info in cell_position_map.items():
                row_start, col_start, row_end, col_end, cell_text = position_info
                
                # 检查是否为合并单元格
                is_merged = row_end > row_start + 1 or col_end > col_start + 1
                
                if is_merged:
                    # 记录合并单元格信息
                    merged_cells.append((row_start, col_start, row_end - 1, col_end - 1))
                    
                    # 标记所有被合并的单元格为已占用
                    for r in range(row_start, row_end):
                        for c in range(col_start, col_end):
                            occupied[r][c] = True
                    
                    # 只在左上角单元格放置内容
                    table_data[row_start][col_start] = cell_text
                else:
                    # 如果单元格未被占用，放置内容
                    if not occupied[row_start][col_start]:
                        table_data[row_start][col_start] = cell_text
            
            return table_data, merged_cells
            
        except Exception as e:
            print(f"构建表格时出错: {e}")
            import traceback
            traceback.print_exc()
            return [], []

    def _detect_merged_cells(self, table):
        """
        检测表格中的合并单元格
        
        参数:
            table: 表格对象
            
        返回:
            合并单元格列表，每个元素为 (行开始, 列开始, 行结束, 列结束)
        """
        merged_cells = []
        
        try:
            # 检查是否是字典对象
            if isinstance(table, dict):
                # 如果是字典对象，直接获取merged_cells字段
                if "merged_cells" in table:
                    return table.get("merged_cells", [])
                    
                # 如果没有merged_cells字段，尝试从cells分析
                if "cells" not in table or not table["cells"]:
                    return []
                    
                cells = table["cells"]
            # 如果是表格对象
            elif hasattr(table, 'cells') and table.cells:
                cells = table.cells
            # 如果表格有tables属性 (PyMuPDF 1.18.0+)
            elif hasattr(table, 'tables') and table.tables:
                # 尝试获取第一个表格
                if len(table.tables) > 0:
                    first_table = table.tables[0]
                    if hasattr(first_table, 'cells') and first_table.cells:
                        cells = first_table.cells
                    else:
                        return []
                else:
                    return []
            else:
                return []
                
            # 收集边界
            rows = set()
            cols = set()
            
            # 处理不同类型的单元格，提取边界信息
            for cell in cells:
                cell_bbox = None
                
                if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                    cell_bbox = cell["bbox"]
                elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                    cell_bbox = cell[:4]
                elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                    cell_bbox = cell.bbox
                
                if not cell_bbox:
                    continue
                    
                rows.add(cell_bbox[1])  # Top
                rows.add(cell_bbox[3])  # Bottom
                cols.add(cell_bbox[0])  # Left
                cols.add(cell_bbox[2])  # Right
                
            # 排序边界
            rows = sorted(rows)
            cols = sorted(cols)
            
            # 映射单元格
            for cell in cells:
                cell_bbox = None
                
                if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                    cell_bbox = cell["bbox"]
                elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                    cell_bbox = cell[:4]
                elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                    cell_bbox = cell.bbox
                
                if not cell_bbox:
                    continue
                
                # 获取索引
                top_idx = rows.index(cell_bbox[1]) if cell_bbox[1] in rows else -1
                bottom_idx = rows.index(cell_bbox[3]) if cell_bbox[3] in rows else -1
                left_idx = cols.index(cell_bbox[0]) if cell_bbox[0] in cols else -1
                right_idx = cols.index(cell_bbox[2]) if cell_bbox[2] in cols else -1
                
                # 检查合并单元格
                if top_idx >= 0 and bottom_idx > top_idx and left_idx >= 0 and right_idx > left_idx:
                    if bottom_idx - top_idx > 1 or right_idx - left_idx > 1:
                        merged_cells.append((top_idx, left_idx, bottom_idx - 1, right_idx - 1))
            
            # 如果上述方法无法检测到合并单元格，尝试备用方法
            if not merged_cells and hasattr(table, 'extract'):
                try:
                    table_data = table.extract()
                    if not table_data:
                        return []
                    
                    rows = len(table_data)
                    if rows == 0:
                        return []
                    
                    cols = len(table_data[0]) if rows > 0 else 0
                    if cols == 0:
                        return []
                    
                    # 跟踪已访问的单元格
                    visited = [[False for _ in range(cols)] for _ in range(rows)]
                    
                    # 检测合并单元格
                    for i in range(rows):
                        for j in range(cols):
                            if visited[i][j]:
                                continue
                            
                            current_value = table_data[i][j]
                            visited[i][j] = True
                            
                            # 检查水平合并
                            col_span = 1
                            for c in range(j + 1, cols):
                                if c < cols and table_data[i][c] == current_value and not visited[i][c]:
                                    col_span += 1
                                    visited[i][c] = True
                                else:
                                    break
                            
                            # 检查垂直合并
                            row_span = 1
                            for r in range(i + 1, rows):
                                if r < rows:
                                    valid_range = True
                                    for c in range(j, min(j + col_span, cols)):
                                        if c >= cols or r >= rows or table_data[r][c] != current_value or visited[r][c]:
                                            valid_range = False
                                            break
                                    
                                    if valid_range:
                                        row_span += 1
                                        for c in range(j, min(j + col_span, cols)):
                                            visited[r][c] = True
                                    else:
                                        break
                                else:
                                    break
                            
                            # 记录合并单元格
                            if row_span > 1 or col_span > 1:
                                merged_cells.append((i, j, i + row_span - 1, j + col_span - 1))
                except Exception as e:
                    print(f"备用合并单元格检测失败: {e}")
        
        except Exception as e:
            print(f"检测合并单元格时出错: {e}")
            import traceback
            traceback.print_exc()
        
        return merged_cells