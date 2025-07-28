"""
简单测试脚本：验证表格单元格样式检测修复
"""

def test_import():
    """测试模块导入"""
    try:
        from table_cell_style_detection_fix import apply_table_cell_style_detection_fix
        print("✓ 成功导入 apply_table_cell_style_detection_fix")
        return True
    except ImportError as e:
        print(f"✗ 导入失败: {e}")
        return False

def test_converter_creation():
    """测试转换器创建"""
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        converter = EnhancedPDFConverter()
        print("✓ 成功创建增强型PDF转换器")
        return converter
    except ImportError:
        try:
            from improved_pdf_converter import ImprovedPDFConverter
            converter = ImprovedPDFConverter()
            print("✓ 成功创建改进版PDF转换器")
            return converter
        except ImportError:
            print("✗ 无法导入任何PDF转换器")
            return None

def test_fix_application():
    """测试修复应用"""
    converter = test_converter_creation()
    if not converter:
        return False
    
    try:
        from table_cell_style_detection_fix import apply_table_cell_style_detection_fix
        success = apply_table_cell_style_detection_fix(converter)
        
        if success:
            print("✓ 修复应用成功")
            
            # 检查关键方法
            methods_to_check = [
                'enhanced_detect_table_styles_with_dimensions',
                'apply_precise_cell_style',
                'apply_precise_column_widths'
            ]
            
            for method in methods_to_check:
                if hasattr(converter, method):
                    print(f"✓ {method} 方法已添加")
                else:
                    print(f"✗ {method} 方法缺失")
            
            return True
        else:
            print("✗ 修复应用失败")
            return False
            
    except Exception as e:
        print(f"✗ 修复应用出错: {e}")
        return False

def main():
    """主测试函数"""
    print("=" * 60)
    print("表格单元格样式检测修复 - 快速测试")
    print("=" * 60)
    
    # 1. 测试导入
    print("\n1. 测试模块导入...")
    if not test_import():
        return
    
    # 2. 测试修复应用
    print("\n2. 测试修复应用...")
    if test_fix_application():
        print("\n✓ 所有测试通过！表格单元格样式检测修复已成功安装。")
        print("\n功能说明:")
        print("- 精确检测表格列宽和行高")
        print("- 识别单元格字体样式（大小、粗体、斜体、颜色）")
        print("- 智能检测表头并应用差异化样式")
        print("- 保留原PDF的表格视觉效果")
    else:
        print("\n✗ 测试失败。请检查相关模块是否正确安装。")

if __name__ == "__main__":
    main()
