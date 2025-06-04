#!/usr/bin/env python
"""
Test script to check module imports
"""

import sys
import os

print(f"Python version: {sys.version}")
print(f"Current directory: {os.getcwd()}")

try:
    import pdf_color_manager
    print("Successfully imported pdf_color_manager")
except Exception as e:
    print(f"Error importing pdf_color_manager: {e}")
    
try:
    import pdf_cmyk_helper
    print("Successfully imported pdf_cmyk_helper")
except Exception as e:
    print(f"Error importing pdf_cmyk_helper: {e}")

try:
    import pdf_font_manager
    print("Successfully imported pdf_font_manager")
except Exception as e:
    print(f"Error importing pdf_font_manager: {e}")

try:
    from enhanced_pdf_converter import EnhancedPDFConverter
    converter = EnhancedPDFConverter()
    print("Successfully initialized EnhancedPDFConverter")
except Exception as e:
    print(f"Error initializing EnhancedPDFConverter: {e}")
    import traceback
    traceback.print_exc()
