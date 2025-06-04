"""
测试增强型表格检测模块
"""

def main():
    """主函数"""
    try:
        # 导入需要的模块
        from enhanced_pdf_converter import EnhancedPDFConverter
        from enhanced_table_detection import apply_enhanced_table_detection_patch
        
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        
        # 应用增强型表格检测补丁
        success = apply_enhanced_table_detection_patch(converter)
        
        # 检查是否成功添加了表格检测方法
        has_detect_tables = hasattr(converter, 'detect_tables')
        has_detect_tables_opencv = hasattr(converter, 'detect_tables_opencv')
        has_detect_tables_by_layout = hasattr(converter, 'detect_tables_by_layout')
        has_detect_tables_by_grid = hasattr(converter, 'detect_tables_by_grid')
        has_detect_tables_by_text_alignment = hasattr(converter, 'detect_tables_by_text_alignment')
        has_analyze_table_structure = hasattr(converter, 'analyze_table_structure')
        
        # 添加提取表格方法
        def _extract_tables(self, pdf_document, page_num):
            """从PDF页面提取表格"""
            try:
                page = pdf_document[page_num]
                
                # 使用增强的表格检测
                if hasattr(self, 'detect_tables'):
                    tables_obj = self.detect_tables(page)
                    if tables_obj and hasattr(tables_obj, 'tables'):
                        return tables_obj.tables
                    else:
                        return []
                
                return []
                
            except Exception as e:
                print(f"表格提取错误: {e}")
                return []
        
        # 添加方法到转换器
        import types
        converter._extract_tables = types.MethodType(_extract_tables, converter)
        
        # 检查是否成功添加了表格提取方法
        has_extract_tables = hasattr(converter, '_extract_tables')
        
        # 输出结果
        print("===== 增强型表格检测测试结果 =====")
        print(f"应用增强型表格检测补丁: {'成功' if success else '失败'}")
        print(f"检测到detect_tables方法: {'是' if has_detect_tables else '否'}")
        print(f"检测到detect_tables_opencv方法: {'是' if has_detect_tables_opencv else '否'}")
        print(f"检测到detect_tables_by_layout方法: {'是' if has_detect_tables_by_layout else '否'}")
        print(f"检测到detect_tables_by_grid方法: {'是' if has_detect_tables_by_grid else '否'}")
        print(f"检测到detect_tables_by_text_alignment方法: {'是' if has_detect_tables_by_text_alignment else '否'}")
        print(f"检测到analyze_table_structure方法: {'是' if has_analyze_table_structure else '否'}")
        print(f"检测到_extract_tables方法: {'是' if has_extract_tables else '否'}")
        print("=================================")
        
        return True
        
    except Exception as e:
        print(f"测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    result = main()
    exit(0 if result else 1)
