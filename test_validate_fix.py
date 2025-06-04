"""
Test script to verify the fixed _validate_and_fix_table_data method
"""
import os
import sys

# Get the current directory to make sure the imports work correctly
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

try:
    # Import the EnhancedPDFConverter class
    from enhanced_pdf_converter import EnhancedPDFConverter
    
    # Create an instance of the converter
    converter = EnhancedPDFConverter()
    
    # Test the _validate_and_fix_table_data method with sample data
    test_table_data = [["A1", "B1"], ["A2", "B2"]]
    test_merged_cells = []
    
    # Call the method that was fixed
    fixed_data, fixed_merged = converter._validate_and_fix_table_data(test_table_data, test_merged_cells)
    
    # Print the results to verify it works
    print("Test successful!")
    print(f"Fixed data: {fixed_data}")
    print(f"Fixed merged cells: {fixed_merged}")
    
    # Write results to a file for verification
    with open("validation_test_result.txt", "w") as f:
        f.write("Test successful!\n")
        f.write(f"Fixed data: {fixed_data}\n")
        f.write(f"Fixed merged cells: {fixed_merged}\n")
    
except Exception as e:
    # If there's an error, write it to a file
    print(f"Error: {e}")
    with open("validation_test_error.txt", "w") as f:
        f.write(f"Error: {e}\n")
        import traceback
        f.write(traceback.format_exc())
