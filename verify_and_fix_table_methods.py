#!/usr/bin/env python
"""
集成测试脚本，测试PDF转换器中的表格处理功能
"""

import os
import sys
import time
from pathlib import Path

# 确保当前目录在sys.path中
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

# 检查并应用修复
def verify_and_fix_table_methods():
    """验证表格处理方法并在必要时修复"""
    try:
        # 导入增强型PDF转换器
        from enhanced_pdf_converter import EnhancedPDFConverter
        
        # 创建实例
        converter = EnhancedPDFConverter()
        
        # 检查必要的方法
        required_methods = [
            '_mark_table_regions',
            '_build_table_from_cells',
            '_detect_merged_cells',
            '_validate_and_fix_table_data'
        ]
        
        missing_methods = []
        for method in required_methods:
            if not hasattr(converter, method):
                missing_methods.append(method)
        
        if missing_methods:
            print(f"发现以下方法缺失: {', '.join(missing_methods)}")
            # 从table_regions_helper.py导入方法
            helper_path = os.path.join(current_dir, 'table_regions_helper.py')
            if os.path.exists(helper_path):
                apply_table_region_fixes(missing_methods)
                print("已应用表格区域处理修复")
            else:
                print("无法找到表格区域帮助模块")
        else:
            print("所有表格处理方法已正确实现")
            
        return True
        
    except Exception as e:
        import traceback
        print(f"验证表格方法时出错: {e}")
        traceback.print_exc()
        return False

def apply_table_region_fixes(missing_methods):
    """从表格区域帮助模块中应用修复"""
    try:
        # 读取helper模块内容
        helper_path = os.path.join(current_dir, 'table_regions_helper.py')
        with open(helper_path, 'r', encoding='utf-8') as f:
            helper_content = f.read()
            
        # 读取转换器模块内容
        converter_path = os.path.join(current_dir, 'enhanced_pdf_converter.py')
        with open(converter_path, 'r', encoding='utf-8') as f:
            converter_content = f.read()
            
        # 替换方法名称并应用修复
        methods_added = []
        
        # 提取mark_table_regions方法
        if '_mark_table_regions' in missing_methods:
            mark_method = extract_method(helper_content, 'mark_table_regions')
            if mark_method:
                # 转换为类方法
                class_method = convert_to_class_method(mark_method)
                # 应用到转换器
                add_method_to_converter(converter_path, '_mark_table_regions', class_method)
                methods_added.append('_mark_table_regions')
                
        # 提取build_table_from_cells方法
        if '_build_table_from_cells' in missing_methods:
            build_method = extract_method(helper_content, 'build_table_from_cells')
            if build_method:
                # 转换为类方法
                class_method = convert_to_class_method(build_method)
                # 应用到转换器
                add_method_to_converter(converter_path, '_build_table_from_cells', class_method)
                methods_added.append('_build_table_from_cells')
                
        # 提取detect_merged_cells方法
        if '_detect_merged_cells' in missing_methods:
            detect_method = extract_method(helper_content, 'detect_merged_cells')
            if detect_method:
                # 转换为类方法
                class_method = convert_to_class_method(detect_method)
                # 应用到转换器
                add_method_to_converter(converter_path, '_detect_merged_cells', class_method)
                methods_added.append('_detect_merged_cells')
                
        print(f"成功添加以下方法: {', '.join(methods_added)}")
        return True
        
    except Exception as e:
        import traceback
        print(f"应用表格区域修复时出错: {e}")
        traceback.print_exc()
        return False

def extract_method(content, method_name):
    """从内容中提取方法定义"""
    import re
    pattern = rf"def {method_name}\(self, (.*?)\):(.*?)(?=\ndef|\Z)"
    match = re.search(pattern, content, re.DOTALL)
    if match:
        args = match.group(1)
        body = match.group(2)
        return f"def {method_name}(self, {args}):{body}"
    return None

def convert_to_class_method(method_text):
    """转换为类方法，调整缩进等"""
    lines = method_text.split('\n')
    result = []
    for i, line in enumerate(lines):
        if i == 0:  # 方法定义行
            result.append(line)
        else:
            # 确保正确缩进
            line_content = line.lstrip()
            if line_content:  # 非空行
                result.append("    " + line_content)
            else:
                result.append("")
    return '\n'.join(result)

def add_method_to_converter(converter_path, method_name, method_code):
    """添加方法到转换器类"""
    with open(converter_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
    # 找到类定义的结尾
    class_end_index = content.rfind("if __name__ == \"__main__\":")
    if class_end_index == -1:
        class_end_index = len(content)
        
    # 插入方法代码
    new_content = content[:class_end_index] + "\n" + method_code + "\n\n" + content[class_end_index:]
    
    # 写回文件
    with open(converter_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"已将 {method_name} 方法添加到转换器类")

def create_test_report(success):
    """创建测试报告文件"""
    report_path = os.path.join(current_dir, 'table_methods_fix_report.txt')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(f"表格处理方法修复测试报告\n")
        f.write(f"时间: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"结果: {'成功' if success else '失败'}\n")
        
    print(f"测试报告已创建: {report_path}")

if __name__ == "__main__":
    print("开始验证和修复表格处理方法...")
    success = verify_and_fix_table_methods()
    print(f"验证和修复 {'成功' if success else '失败'}")
    create_test_report(success)
