#!/usr/bin/env python
"""
PDF转Word增强转换工具
此工具应用了所有的增强修复，提供命令行接口用于PDF转Word转换
"""

import os
import sys
import argparse
import traceback
from pathlib import Path

def main():
    """主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description="增强型PDF转Word转换工具")
    parser.add_argument("input", help="输入PDF文件路径")
    parser.add_argument("output", nargs="?", help="输出Word文件路径 (可选，默认使用相同文件名但扩展名为.docx)")
    parser.add_argument("--format", "-f", choices=["docx", "excel"], default="docx", help="输出格式，docx或excel (默认为docx)")
    parser.add_argument("--gui", "-g", action="store_true", help="启动GUI界面而不是命令行转换")
    args = parser.parse_args()
    
    # 如果选择了GUI模式，启动GUI
    if args.gui:
        run_gui()
        return
    
    # 验证输入文件
    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f"错误: 输入文件 {input_path} 不存在")
        return
    
    # 确定输出文件路径
    if args.output:
        output_path = os.path.abspath(args.output)
    else:
        # 默认使用相同文件名但扩展名为.docx或.xlsx
        if args.format == "docx":
            output_path = os.path.splitext(input_path)[0] + ".docx"
        else:
            output_path = os.path.splitext(input_path)[0] + ".xlsx"
    
    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)
    
    # 导入必要的模块
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        from all_pdf_fixes_integrator import integrate_all_fixes
    except ImportError as e:
        print(f"错误: 导入必要模块失败: {e}")
        print("请确保已安装所有必要的依赖")
        return
    
    print(f"正在转换 {input_path} 为 {args.format.upper()}...")
    
    # 创建转换器实例
    converter = EnhancedPDFConverter()
    
    # 应用所有修复
    try:
        converter = integrate_all_fixes(converter)
        print("已应用所有PDF转换修复")
    except Exception as e:
        print(f"警告: 应用修复时出错: {e}")
        traceback.print_exc()
        print("将尝试继续使用未修复的转换器")
    
    # 执行转换
    try:
        if args.format == "docx":
            result = converter.convert_pdf_to_docx(input_path, output_path)
            print(f"转换完成: {result}")
        else:
            result = converter.convert_pdf_to_excel(input_path, output_path)
            print(f"转换完成: {result}")
            
        # 检查转换结果
        if os.path.exists(result):
            print(f"成功! 输出文件: {result}")
            print(f"文件大小: {os.path.getsize(result) / 1024:.2f} KB")
        else:
            print(f"错误: 转换可能失败，找不到输出文件: {result}")
    except Exception as e:
        print(f"转换过程中出错: {e}")
        traceback.print_exc()

def run_gui():
    """启动GUI界面"""
    try:
        import run_gui_with_all_fixes
        run_gui_with_all_fixes.main()
    except ImportError:
        print("错误: 无法导入GUI模块")
        print("请确保run_gui_with_all_fixes.py文件存在")
    except Exception as e:
        print(f"启动GUI时出错: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
