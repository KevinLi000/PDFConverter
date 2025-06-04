"""
测试增强型表格检测功能
"""

import os
import sys
import argparse
import traceback

def test_enhanced_table_detection(pdf_path):
    """
    测试增强型表格检测
    
    参数:
        pdf_path: PDF文件路径
    """
    try:
        # 导入增强型PDF转换器
        from enhanced_pdf_converter import EnhancedPDFConverter
        
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        
        # 设置PDF路径
        converter.pdf_path = pdf_path
        
        # 设置高分辨率以提高表格检测精度
        converter.dpi = 400
        
        # 应用增强型表格检测补丁
        from enhanced_table_detection import apply_enhanced_table_detection_patch
        apply_enhanced_table_detection_patch(converter)
        
        print(f"正在检测PDF文件中的表格: {pdf_path}")
        
        # 打开PDF
        import fitz
        pdf_doc = fitz.open(pdf_path)
        
        # 检测每一页的表格
        total_tables = 0
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc[page_num]
            print(f"\n处理第 {page_num + 1} 页...")
            
            # 使用增强的表格检测
            tables = converter.detect_tables(page)
            
            if hasattr(tables, 'tables'):
                page_tables = tables.tables
                page_table_count = len(page_tables)
                print(f"在第 {page_num + 1} 页检测到 {page_table_count} 个表格")
                total_tables += page_table_count
                
                # 输出表格信息
                for i, table in enumerate(page_tables):
                    print(f"  表格 {i+1}: 位置 = {table.get('bbox', 'N/A')}")
                    print(f"    行数 = {len(table.get('rows', []))}")
                    print(f"    列数 = {len(table.get('cols', []))}")
            else:
                print(f"在第 {page_num + 1} 页未检测到表格")
        
        # 关闭PDF
        pdf_doc.close()
        
        print(f"\n总计检测到 {total_tables} 个表格")
        
        return total_tables
    
    except Exception as e:
        print(f"测试失败: {e}")
        traceback.print_exc()
        return 0

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="测试增强型表格检测功能")
    parser.add_argument("pdf_path", help="要检测表格的PDF文件路径")
    args = parser.parse_args()
    
    if not os.path.exists(args.pdf_path):
        print(f"错误: PDF文件不存在: {args.pdf_path}")
        sys.exit(1)
    
    tables_count = test_enhanced_table_detection(args.pdf_path)
    
    if tables_count > 0:
        print(f"测试成功! 检测到 {tables_count} 个表格")
        sys.exit(0)
    else:
        print("测试失败! 未检测到任何表格")
        sys.exit(1)

if __name__ == "__main__":
    main()
