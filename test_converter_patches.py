#!/usr/bin/env python
"""
测试PDF转换器补丁功能
"""

from enhanced_pdf_converter import EnhancedPDFConverter
from improved_pdf_converter import ImprovedPDFConverter
from converter_patches import apply_converter_patches

def test_converter_patches():
    """测试转换器补丁是否正常工作"""
    # 创建两种类型的转换器
    enhanced = EnhancedPDFConverter()
    improved = ImprovedPDFConverter()
    
    # 检查补丁前的方法
    print("补丁前 EnhancedPDFConverter 的方法:")
    has_method_before = hasattr(enhanced, '_process_text_block_enhanced')
    print(f"_process_text_block_enhanced 方法存在: {has_method_before}")
    
    # 应用补丁
    print("\n应用补丁...")
    apply_converter_patches(enhanced)
    
    # 检查补丁后的方法
    print("\n补丁后 EnhancedPDFConverter 的方法:")
    has_method_after = hasattr(enhanced, '_process_text_block_enhanced')
    print(f"_process_text_block_enhanced 方法存在: {has_method_after}")
    
    # 检查补丁是否对其他转换器也生效
    apply_converter_patches(improved)
    has_method_improved = hasattr(improved, '_process_text_block_enhanced')
    print(f"ImprovedPDFConverter 中 _process_text_block_enhanced 方法存在: {has_method_improved}")
    
    # 验证方法可以被调用
    if has_method_after:
        print("\n验证方法可以被调用...")
        try:
            # 创建一个简单的块和段落进行测试
            from docx import Document
            doc = Document()
            paragraph = doc.add_paragraph()
            block = {"lines": [{"spans": [{"text": "测试文本", "font": "Arial", "size": 12}]}]}
            
            # 调用方法
            enhanced._process_text_block_enhanced(paragraph, block)
            print("方法调用成功!")
        except Exception as e:
            print(f"方法调用失败: {e}")
    
    return has_method_before, has_method_after

if __name__ == "__main__":
    print("开始测试PDF转换器补丁...")
    before, after = test_converter_patches()
    
    if not before and after:
        print("\n测试结果: 补丁成功添加了缺失的方法!")
    else:
        print("\n测试结果: 补丁未能正确添加方法或方法已存在!")
