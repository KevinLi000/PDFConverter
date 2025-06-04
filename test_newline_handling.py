#!/usr/bin/env python
"""
测试PDF转Word中的换行符处理
"""

import os
import sys
import tempfile
import argparse
import traceback
from pathlib import Path

# 确保必要的模块可用
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

def test_newline_handling(pdf_path, output_dir=None):
    """测试换行符处理功能"""
    print("\n===== 测试PDF转Word换行符处理 =====")
    
    try:
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        
        # 应用所有修复
        converter = integrate_all_fixes(converter)
        
        # 设置路径
        converter.set_paths(pdf_path, output_dir)
        
        # 启用最大格式保留模式
        converter.enhance_format_preservation()
        
        # 执行转换，使用高级模式
        output_path = converter.pdf_to_word(method="advanced")
        
        print(f"成功将PDF转换为Word文档: {output_path}")
        print("请检查Word文档中的表格单元格是否正确处理了换行符")
        return True
        
    except Exception as e:
        print(f"换行符处理测试失败: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    # 解析命令行参数
    parser = argparse.ArgumentParser(description="测试PDF转Word换行符处理")
    parser.add_argument("pdf_path", help="要处理的PDF文件路径")
    parser.add_argument("--output_dir", "-o", help="输出目录，默认为PDF所在目录")
    args = parser.parse_args()
    
    # 运行测试
    test_newline_handling(args.pdf_path, args.output_dir)
