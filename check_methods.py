import sys
import os

def check_methods():
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        converter = EnhancedPDFConverter()
        
        methods = [
            '_mark_table_regions',
            '_build_table_from_cells',
            '_detect_merged_cells',
            '_validate_and_fix_table_data'
        ]
        
        result = "Table methods test results:\n"
        for method in methods:
            has_method = hasattr(converter, method)
            result += f"{method}: {'✓' if has_method else '✗'}\n"
        
        # Write to file
        with open("method_check_result.txt", "w") as f:
            f.write(result)
        
        return True
    except Exception as e:
        with open("method_check_error.txt", "w") as f:
            f.write(f"Error: {str(e)}")
        return False

if __name__ == "__main__":
    check_methods()
