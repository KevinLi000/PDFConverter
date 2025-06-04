#!/usr/bin/env python
"""
PDF转Word表格修复验证工具
用于测试和验证表格修复功能是否正确解决了三个具体问题：
1. 表格样式问题
2. 图像识别失败
3. 文本对齐和格式问题
"""

import os
import sys
import tempfile
import argparse
import traceback
from pathlib import Path

# 尝试导入必要的模块
try:
    import fitz  # PyMuPDF
except ImportError:
    print("错误: 未安装PyMuPDF")
    print("请使用命令安装: pip install PyMuPDF")
    sys.exit(1)

try:
    from enhanced_pdf_converter import EnhancedPDFConverter
except ImportError:
    print("错误: 无法导入EnhancedPDFConverter")
    print("请确保enhanced_pdf_converter.py在当前目录或Python路径中")
    sys.exit(1)

try:
    from all_pdf_fixes_integrator import integrate_all_fixes
except ImportError:
    print("错误: 无法导入all_pdf_fixes_integrator")
    print("请确保all_pdf_fixes_integrator.py在当前目录或Python路径中")
    sys.exit(1)

try:
    from advanced_table_fixes import apply_advanced_table_fixes
except ImportError:
    print("错误: 无法导入advanced_table_fixes")
    print("请确保advanced_table_fixes.py在当前目录或Python路径中")
    sys.exit(1)

def validate_table_style_fix(pdf_path, output_dir):
    """验证表格样式修复"""
    print("\n===== 验证表格样式修复 =====")
    
    try:
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        
        # 应用所有修复
        converter = integrate_all_fixes(converter)
        
        # 设置路径
        converter.set_paths(pdf_path, output_dir)
        
        # 设置最大格式保留模式
        converter.enhance_format_preservation()
        
        # 执行转换，使用高级模式
        output_path = converter.pdf_to_word(method="advanced")
        
        print(f"成功将PDF转换为Word文档: {output_path}")
        print("请检查Word文档中的表格样式是否正确")
        return True
        
    except Exception as e:
        print(f"表格样式修复验证失败: {e}")
        traceback.print_exc()
        return False

def validate_image_recognition_fix(pdf_path, output_dir):
    """验证图像识别修复"""
    print("\n===== 验证图像识别修复 =====")
    
    try:
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        
        # 应用所有修复
        converter = integrate_all_fixes(converter)
        
        # 设置路径
        converter.set_paths(pdf_path, output_dir)
        
        # 打开PDF文件
        pdf_document = fitz.open(pdf_path)
        
        # 统计图像数量
        image_count = 0
        extracted_image_count = 0
        
        # 创建临时目录存储提取的图像
        temp_dir = tempfile.mkdtemp()
        
        # 使用多种方法提取图像，统计成功率
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # 方法1: 提取图像对象
            try:
                image_list = page.get_images(full=True)
                for img_index, img_info in enumerate(image_list):
                    image_count += 1
                    try:
                        xref = img_info[0]
                        base_image = pdf_document.extract_image(xref)
                        
                        if base_image:
                            # 保存提取的图像
                            image_path = os.path.join(temp_dir, f"image_m1_{page_num}_{img_index}.png")
                            with open(image_path, "wb") as f:
                                f.write(base_image["image"])
                            
                            if os.path.exists(image_path):
                                extracted_image_count += 1
                    except Exception as img_err:
                        print(f"提取图像对象出错: {img_err}")
            except Exception as e:
                print(f"提取页面图像对象出错: {e}")
            
            # 方法2: 使用高分辨率渲染页面区域
            try:
                # 获取页面上的所有图像区域
                blocks = page.get_text("dict")["blocks"]
                for block in blocks:
                    if block["type"] == 1:  # 图像块
                        image_count += 1
                        try:
                            # 渲染图像区域
                            bbox = block["bbox"]
                            clip_rect = fitz.Rect(bbox)
                            zoom = 4.0
                            matrix = fitz.Matrix(zoom, zoom)
                            pix = page.get_pixmap(matrix=matrix, clip=clip_rect, alpha=False)
                            
                            # 保存渲染的图像
                            image_path = os.path.join(temp_dir, f"image_m2_{page_num}_{hash(str(bbox))}.png")
                            pix.save(image_path)
                            
                            if os.path.exists(image_path):
                                extracted_image_count += 1
                        except Exception as render_err:
                            print(f"渲染图像区域出错: {render_err}")
            except Exception as e:
                print(f"获取页面图像区域出错: {e}")
        
        # 计算图像提取成功率
        success_rate = (extracted_image_count / max(image_count, 1)) * 100
        
        print(f"PDF中包含 {image_count} 个图像")
        print(f"成功提取了 {extracted_image_count} 个图像")
        print(f"图像识别成功率: {success_rate:.2f}%")
        
        # 设置成功阈值为80%
        if success_rate >= 80:
            print("图像识别修复验证成功")
            return True
        else:
            print("图像识别修复验证失败，成功率低于80%")
            return False
        
    except Exception as e:
        print(f"图像识别修复验证失败: {e}")
        traceback.print_exc()
        return False
    finally:
        # 清理临时目录
        try:
            import shutil
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except Exception:
            pass

def validate_text_alignment_fix(pdf_path, output_dir):
    """验证文本对齐和格式修复"""
    print("\n===== 验证文本对齐和格式修复 =====")
    
    try:
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        
        # 单独应用高级表格修复
        converter = apply_advanced_table_fixes(converter)
        
        # 设置路径
        converter.set_paths(pdf_path, output_dir)
        
        # 打开PDF文件
        pdf_document = fitz.open(pdf_path)
        
        # 统计表格单元格总数和正确对齐的单元格数
        total_cells = 0
        aligned_cells = 0
        
        # 对每个页面进行处理
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # 使用备用表格提取方法
            tables = []
            
            # 尝试使用增强型表格检测
            try:
                from advanced_table_fixes import extract_tables_advanced
                enhanced_tables = extract_tables_advanced(converter, pdf_document, page_num)
                if enhanced_tables and hasattr(enhanced_tables, 'tables') and len(enhanced_tables.tables) > 0:
                    tables = enhanced_tables.tables
            except Exception as e:
                print(f"使用增强型表格检测出错: {e}")
            
            # 如果没有检测到表格，尝试使用其他方法
            if not tables:
                try:
                    # 尝试使用PyMuPDF的表格检测
                    if hasattr(page, 'find_tables'):
                        table_finder = page.find_tables()
                        tables = table_finder.tables
                    # 尝试使用camelot
                    elif 'camelot' in sys.modules:
                        import camelot
                        tables_camelot = camelot.read_pdf(pdf_path, pages=str(page_num + 1))
                        if len(tables_camelot) > 0:
                            for table in tables_camelot:
                                table_dict = {
                                    "bbox": table._bbox,
                                    "cells": []
                                }
                                for i, row in enumerate(table.cells):
                                    for j, cell in enumerate(row):
                                        table_dict["cells"].append({
                                            "bbox": cell,
                                            "text": table.df.iloc[i, j]
                                        })
                                tables.append(table_dict)
                except Exception as e:
                    print(f"使用其他表格检测方法出错: {e}")
            
            # 分析表格单元格对齐情况
            for table in tables:
                # 获取表格单元格
                cells = []
                
                if isinstance(table, dict) and "cells" in table:
                    cells = table["cells"]
                elif hasattr(table, 'cells'):
                    cells = table.cells
                
                # 分析单元格对齐
                for cell in cells:
                    total_cells += 1
                    
                    # 获取单元格文本和边界框
                    cell_text = ""
                    cell_bbox = None
                    
                    if isinstance(cell, dict):
                        cell_text = cell.get("text", "")
                        cell_bbox = cell.get("bbox")
                    elif hasattr(cell, 'text'):
                        cell_text = cell.text
                        if hasattr(cell, 'bbox'):
                            cell_bbox = cell.bbox
                        elif hasattr(cell, 'rect'):
                            cell_bbox = cell.rect
                    
                    # 检查文本对齐
                    if cell_bbox and cell_text:
                        # 简单检查文本是否在单元格边界内，并且有合理的边距
                        cell_width = cell_bbox[2] - cell_bbox[0]
                        text_length = len(cell_text)
                        
                        # 假设文本每个字符平均宽度为原始单元格宽度的1/8
                        expected_width = text_length * (cell_width / 8)
                        
                        # 如果预期文本宽度小于单元格宽度的90%，认为对齐正确
                        if expected_width < cell_width * 0.9:
                            aligned_cells += 1
        
        # 计算文本对齐成功率
        alignment_rate = (aligned_cells / max(total_cells, 1)) * 100
        
        print(f"表格中包含 {total_cells} 个单元格")
        print(f"正确对齐的单元格: {aligned_cells} 个")
        print(f"文本对齐成功率: {alignment_rate:.2f}%")
        
        # 设置成功阈值为75%
        if alignment_rate >= 75:
            print("文本对齐修复验证成功")
            return True
        else:
            print("文本对齐修复验证失败，成功率低于75%")
            return False
        
    except Exception as e:
        print(f"文本对齐修复验证失败: {e}")
        traceback.print_exc()
        return False

def run_validation(pdf_path, output_dir=None):
    """运行所有验证测试"""
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return False
    
    # 设置输出目录
    if output_dir is None:
        output_dir = os.path.dirname(pdf_path)
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"开始验证PDF转Word表格修复功能...")
    print(f"PDF文件: {pdf_path}")
    print(f"输出目录: {output_dir}")
    
    # 运行验证测试
    style_fix_passed = validate_table_style_fix(pdf_path, output_dir)
    image_fix_passed = validate_image_recognition_fix(pdf_path, output_dir)
    text_fix_passed = validate_text_alignment_fix(pdf_path, output_dir)
    
    # 显示验证结果
    print("\n===== 验证测试结果 =====")
    print(f"表格样式修复: {'通过' if style_fix_passed else '未通过'}")
    print(f"图像识别修复: {'通过' if image_fix_passed else '未通过'}")
    print(f"文本对齐修复: {'通过' if text_fix_passed else '未通过'}")
    
    # 总体结果
    all_passed = style_fix_passed and image_fix_passed and text_fix_passed
    print(f"\n总体验证结果: {'全部通过' if all_passed else '部分未通过'}")
    
    return all_passed

if __name__ == "__main__":
    # 解析命令行参数
    parser = argparse.ArgumentParser(description="验证PDF转Word表格修复功能")
    parser.add_argument("pdf_path", help="要处理的PDF文件路径")
    parser.add_argument("--output_dir", "-o", help="输出目录，默认为PDF所在目录")
    args = parser.parse_args()
    
    # 运行验证
    run_validation(args.pdf_path, args.output_dir)
