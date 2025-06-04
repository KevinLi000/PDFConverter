"""
应用增强型表格检测修复
"""
import os
import sys
import types

def apply_table_fixes():
    """应用表格检测修复到增强型PDF转换器"""
    try:
        # 导入增强型PDF转换器
        from enhanced_pdf_converter import EnhancedPDFConverter
        
        # 导入增强型表格检测模块
        from enhanced_table_detection import apply_enhanced_table_detection_patch
        
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        
        # 应用增强型表格检测补丁
        apply_enhanced_table_detection_patch(converter)
        
        # 添加提取表格方法
        def extract_tables(self, pdf_document, page_num):
            """
            从PDF页面提取表格
            
            参数:
                pdf_document: PDF文档对象
                page_num: 页码
                
            返回:
                检测到的表格列表
            """
            try:
                page = pdf_document[page_num]
                
                # 使用增强的表格检测
                if hasattr(self, 'detect_tables'):
                    # 使用增强的detect_tables方法
                    tables_obj = self.detect_tables(page)
                    if tables_obj and hasattr(tables_obj, 'tables'):
                        return tables_obj.tables
                    else:
                        return []
                
                # 备用方法：尝试使用find_tables (如果可用)
                try:
                    tables = page.find_tables()
                    if tables and len(tables.tables) > 0:
                        return tables.tables
                except (AttributeError, TypeError) as e:
                    print(f"表格检测警告: {e}")
                
                return []
                
            except Exception as e:
                print(f"表格提取错误: {e}")
                import traceback
                traceback.print_exc()
                return []
                
        # 绑定方法到转换器类
        EnhancedPDFConverter._extract_tables = extract_tables
        
        print("已添加增强型表格提取方法到PDF转换器")
        return True
        
    except ImportError as e:
        print(f"导入增强型PDF转换器失败: {e}")
        return False

# 如果作为脚本执行，应用修复
if __name__ == "__main__":
    success = apply_table_fixes()
    if success:
        print("成功应用表格修复")
    else:
        print("应用表格修复失败")
