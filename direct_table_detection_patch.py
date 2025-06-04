"""
直接添加表格检测功能到PDF转换器，解决find_tables兼容性问题
"""

import types
import os
import tempfile
import traceback

def patch_table_detection(converter):
    """
    直接为转换器添加表格检测能力，
    不依赖于PyMuPDF的find_tables方法
    
    参数:
        converter: EnhancedPDFConverter或ImprovedPDFConverter实例
    
    返回:
        bool: 是否成功应用补丁
    """
    try:
        import fitz
        
        # 如果没有detect_tables方法，添加一个
        if not hasattr(converter, 'detect_tables'):
            # 首先尝试从备用模块导入
            try:
                from table_detection_backup import extract_tables_opencv
                
                def detect_tables_wrapper(self, page):
                    """使用OpenCV进行表格检测的包装方法"""
                    return extract_tables_opencv(page, dpi=self.dpi if hasattr(self, 'dpi') else 300)
                
                converter.detect_tables = types.MethodType(detect_tables_wrapper, converter)
                print("使用OpenCV表格检测作为备用")
            except ImportError:
                # 如果备用模块不可用，实现内联版本
                try:
                    import cv2
                    import numpy as np
                    from PIL import Image
                    
                    def detect_tables_inline(self, page):
                        """内联实现的表格检测方法"""
                        try:
                            # 渲染页面为图像
                            zoom = 2  # 放大因子，提高分辨率
                            mat = fitz.Matrix(zoom, zoom)
                            pix = page.get_pixmap(matrix=mat)
                            
                            # 转换为PIL图像
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            
                            # 转换为OpenCV格式
                            img_np = np.array(img)
                            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
                            
                            # 使用自适应阈值处理
                            binary = cv2.adaptiveThreshold(
                                gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 2
                            )
                            
                            # 查找水平和垂直线
                            horizontal = binary.copy()
                            vertical = binary.copy()
                            
                            # 处理水平线
                            horizontalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
                            horizontal = cv2.morphologyEx(horizontal, cv2.MORPH_OPEN, horizontalStructure)
                            
                            # 处理垂直线
                            verticalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
                            vertical = cv2.morphologyEx(vertical, cv2.MORPH_OPEN, verticalStructure)
                            
                            # 合并水平线和垂直线
                            table_mask = cv2.bitwise_or(horizontal, vertical)
                            
                            # 查找轮廓 - 这些是潜在的表格
                            contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                            
                            # 提取表格区域
                            tables = []
                            page_width, page_height = page.rect.width, page.rect.height
                            scale_x, scale_y = page_width / pix.width, page_height / pix.height
                            
                            for contour in contours:
                                x, y, w, h = cv2.boundingRect(contour)
                                # 过滤掉太小的区域
                                if w > 100 and h > 100:
                                    # 转换回PDF坐标系
                                    pdf_x0 = x * scale_x
                                    pdf_y0 = y * scale_y
                                    pdf_x1 = (x + w) * scale_x
                                    pdf_y1 = (y + h) * scale_y
                                    
                                    # 添加到结果
                                    tables.append({
                                        "bbox": (pdf_x0, pdf_y0, pdf_x1, pdf_y1),
                                        "type": "table"
                                    })
                            
                            return tables
                            
                        except Exception as e:
                            print(f"表格检测错误: {e}")
                            return []
                    
                    converter.detect_tables = types.MethodType(detect_tables_inline, converter)
                    print("使用内联表格检测方法")
                except ImportError:
                    print("警告: 无法导入OpenCV，表格检测功能将受限")
                    
                    # 提供一个最小化的表格检测方法
                    def minimal_table_detection(self, page):
                        """最小化的表格检测，基于文本块分析"""
                        tables = []
                        try:
                            # 获取页面文本块
                            blocks = page.get_text("dict")["blocks"]
                            
                            # 按类型过滤块
                            text_blocks = [b for b in blocks if b["type"] == 0]
                            
                            # 如果块太少，可能没有表格
                            if len(text_blocks) < 4:
                                return []
                                
                            # 简单分析：连续排列的小文本块可能是表格的单元格
                            # 排序块
                            text_blocks.sort(key=lambda b: (b["bbox"][1], b["bbox"][0]))
                            
                            # 检测可能的表格行
                            y_positions = {}
                            for block in text_blocks:
                                y = int(block["bbox"][1])
                                if y in y_positions:
                                    y_positions[y].append(block)
                                else:
                                    y_positions[y] = [block]
                            
                            # 如果有多行具有类似的x坐标，可能是表格
                            if len([row for row in y_positions.values() if len(row) > 1]) >= 2:
                                # 粗略估计表格范围
                                all_blocks = [b for row in y_positions.values() if len(row) > 1 for b in row]
                                if all_blocks:
                                    min_x = min(b["bbox"][0] for b in all_blocks)
                                    min_y = min(b["bbox"][1] for b in all_blocks)
                                    max_x = max(b["bbox"][2] for b in all_blocks)
                                    max_y = max(b["bbox"][3] for b in all_blocks)
                                    
                                    # 添加到结果
                                    tables.append({
                                        "bbox": (min_x, min_y, max_x, max_y),
                                        "type": "table"
                                    })
                            
                            return tables
                        except Exception as e:
                            print(f"简单表格检测错误: {e}")
                            return []
                    
                    converter.detect_tables = types.MethodType(minimal_table_detection, converter)
                    print("使用最小化表格检测方法")
        
        # 确保有_extract_tables方法
        if not hasattr(converter, '_extract_tables'):
            def extract_tables_from_detect(self, pdf_document, page_num):
                """从detect_tables方法提取表格"""
                try:
                    page = pdf_document[page_num]
                    if hasattr(self, 'detect_tables'):
                        return self.detect_tables(page)
                    return []
                except Exception as e:
                    print(f"表格提取错误: {e}")
                    return []
            
            converter._extract_tables = types.MethodType(extract_tables_from_detect, converter)
        
        return True
    
    except Exception as e:
        print(f"应用表格检测补丁错误: {e}")
        traceback.print_exc()
        return False

# 测试代码
if __name__ == "__main__":
    print("表格检测补丁模块 - 可以导入到转换应用程序中使用")
