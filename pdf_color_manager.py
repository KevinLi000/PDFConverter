#!/usr/bin/env python
"""
PDF颜色管理模块
用于处理PDF颜色空间转换和颜色增强
"""

import sys
import traceback

# 尝试导入需要的库
try:
    import fitz  # PyMuPDF
except ImportError:
    print("错误: 未安装PyMuPDF")
    print("请使用命令安装: pip install PyMuPDF")

try:
    from PIL import Image, ImageEnhance
except ImportError:
    print("错误: 未安装Pillow")
    print("请使用命令安装: pip install Pillow")

class PDFColorManager:
    """PDF颜色管理器，处理不同颜色空间和颜色增强"""
    
    def __init__(self):
        """初始化颜色管理器"""
        self.color_enhancement_level = 1.0  # 颜色增强级别
        self.contrast_enhancement_level = 1.0  # 对比度增强级别
        self.saturation_enhancement_level = 1.0  # 饱和度增强级别
        self.brightness_adjustment = 0.0  # 亮度调整
        self.color_balance = [1.0, 1.0, 1.0]  # RGB颜色平衡
        self.use_color_profiles = False  # 是否使用ICC颜色配置文件
        self.icc_profile_path = None  # ICC配置文件路径
    
    def set_enhancement_levels(self, color=None, contrast=None, saturation=None, brightness=None):
        """
        设置各种增强级别
        
        参数:
            color: 颜色增强级别 (1.0为正常)
            contrast: 对比度增强级别 (1.0为正常)
            saturation: 饱和度增强级别 (1.0为正常)
            brightness: 亮度调整 (0.0为正常)
        """
        if color is not None:
            self.color_enhancement_level = max(0.5, min(1.5, float(color)))
        if contrast is not None:
            self.contrast_enhancement_level = max(0.5, min(1.5, float(contrast)))
        if saturation is not None:
            self.saturation_enhancement_level = max(0.5, min(1.5, float(saturation)))
        if brightness is not None:
            self.brightness_adjustment = max(-0.5, min(0.5, float(brightness)))
    
    def set_color_balance(self, red=1.0, green=1.0, blue=1.0):
        """
        设置RGB颜色平衡
        
        参数:
            red: 红色分量调整 (1.0为正常)
            green: 绿色分量调整 (1.0为正常)
            blue: 蓝色分量调整 (1.0为正常)
        """
        self.color_balance = [
            max(0.5, min(1.5, float(red))),
            max(0.5, min(1.5, float(green))),
            max(0.5, min(1.5, float(blue)))
        ]
    
    def set_icc_profile(self, profile_path):
        """
        设置ICC颜色配置文件
        
        参数:
            profile_path: ICC配置文件的路径
        """
        self.icc_profile_path = profile_path
        self.use_color_profiles = True
    
    def enhance_image(self, image):
        """
        增强图像颜色
        
        参数:
            image: PIL.Image对象
            
        返回:
            增强后的图像
        """
        try:
            # 应用对比度增强
            if self.contrast_enhancement_level != 1.0:
                enhancer = ImageEnhance.Contrast(image)
                image = enhancer.enhance(self.contrast_enhancement_level)
            
            # 应用饱和度增强
            if self.saturation_enhancement_level != 1.0:
                enhancer = ImageEnhance.Color(image)
                image = enhancer.enhance(self.saturation_enhancement_level)
            
            # 应用亮度调整
            if self.brightness_adjustment != 0.0:
                enhancer = ImageEnhance.Brightness(image)
                factor = 1.0 + self.brightness_adjustment
                image = enhancer.enhance(factor)
            
            # 应用颜色平衡（需要更复杂的处理）
            if self.color_balance != [1.0, 1.0, 1.0]:
                r, g, b = self.color_balance
                if image.mode == 'RGB':
                    r_band, g_band, b_band = image.split()
                    r_band = ImageEnhance.Brightness(r_band).enhance(r)
                    g_band = ImageEnhance.Brightness(g_band).enhance(g)
                    b_band = ImageEnhance.Brightness(b_band).enhance(b)
                    image = Image.merge('RGB', (r_band, g_band, b_band))
            
            return image
            
        except Exception as e:
            print(f"图像增强出错: {e}")
            traceback.print_exc()
            return image  # 返回原始图像
    
    def handle_pixmap_color(self, pixmap):
        """
        处理PyMuPDF Pixmap的颜色空间转换和增强
        
        参数:
            pixmap: fitz.Pixmap对象
            
        返回:
            处理后的pixmap
        """
        try:
            # 检查颜色空间
            if hasattr(pixmap, 'colorspace') and pixmap.colorspace:
                cs_name = pixmap.colorspace.name if hasattr(pixmap.colorspace, 'name') else str(pixmap.colorspace)
                
                # CMYK颜色空间转换为RGB
                if cs_name in ("CMYK", "DeviceCMYK"):
                    pixmap = fitz.Pixmap(fitz.csRGB, pixmap)
                
                # 处理其他特殊颜色空间
                elif pixmap.n > 4:  # 如果通道数 > 4，可能是其他复杂颜色空间
                    # 创建RGB颜色空间的pixmap
                    pixmap = fitz.Pixmap(fitz.csRGB, pixmap)
            
            return pixmap
            
        except Exception as e:
            print(f"处理Pixmap颜色空间时出错: {e}")
            traceback.print_exc()
            return pixmap  # 返回原始pixmap

def handle_pixmap_color(pixmap):
    """
    全局函数用于处理pixmap颜色，方便直接调用
    
    参数:
        pixmap: fitz.Pixmap对象
        
    返回:
        处理后的pixmap
    """
    manager = PDFColorManager()
    return manager.handle_pixmap_color(pixmap)
