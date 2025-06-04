"""
PYTHON-DOCX TABLE MERGE PATCH

This module provides a permanent patch for the python-docx library to fix the
"requested span not rectangular" error that occurs when merging table cells.

The error occurs in the _span_dimensions method of the CT_Tc class, which validates
that cells being merged form a perfect rectangle. This module replaces that method
with a more permissive version.

Usage:
    from docx_table_merge_patch import apply_patch, remove_patch

    # Apply the patch at the start of your application
    apply_patch()

    # Your code using python-docx here...
    
    # Optionally, remove the patch when done
    remove_patch()
"""

import sys
import os
import warnings
from importlib import import_module
import inspect
import types

# Information about the target function to patch
PACKAGE_NAME = 'docx'
MODULE_PATH = 'docx.oxml.table'
CLASS_NAME = 'CT_Tc'
METHOD_NAME = '_span_dimensions'

# Original method reference (will be set when patch is applied)
original_method = None

# The replacement method implementation
def patched_span_dimensions(self, other_tc):
    """
    A patched version of _span_dimensions that bypasses the rectangular validation
    and simply calculates the dimensions needed to encompass all selected cells.
    
    Args:
        self: The source cell
        other_tc: The target cell to merge with
        
    Returns:
        A tuple of (top, left, height, width) for the merged cell
    """
    # Calculate the dimensions without validation
    top = min(self.top, other_tc.top)
    left = min(self.left, other_tc.left)
    bottom = max(self.bottom, other_tc.bottom)
    right = max(self.right, other_tc.right)
    
    return top, left, bottom - top, right - left

def get_module_and_class():
    """
    Get the module and class to patch.
    
    Returns:
        tuple: (module, class_object) or (None, None) if not found
    """
    try:
        # Import the module containing the class
        module = import_module(MODULE_PATH)
        
        # Get the class object
        class_object = getattr(module, CLASS_NAME)
        
        return module, class_object
    except (ImportError, AttributeError) as e:
        print(f"Error accessing {MODULE_PATH}.{CLASS_NAME}: {e}")
        return None, None

def apply_patch(verbose=True):
    """
    Apply the patch to fix the table cell merging issue.
    
    Args:
        verbose: Whether to print information about the patch
        
    Returns:
        bool: True if patch was applied successfully, False otherwise
    """
    global original_method
    
    # Check if the patch is already applied
    if original_method is not None:
        if verbose:
            print("Patch is already applied.")
        return True
    
    # Get the module and class
    module, class_object = get_module_and_class()
    if module is None or class_object is None:
        return False
    
    try:
        # Get the original method
        original_method = getattr(class_object, METHOD_NAME)
        
        # Create a new method with our implementation
        new_method = types.MethodType(patched_span_dimensions, None)
        
        # Replace the method in the class
        setattr(class_object, METHOD_NAME, new_method.__func__)
        
        if verbose:
            print(f"Successfully patched {MODULE_PATH}.{CLASS_NAME}.{METHOD_NAME}")
        
        return True
    except Exception as e:
        print(f"Error applying patch: {e}")
        original_method = None
        return False

def remove_patch(verbose=True):
    """
    Remove the patch and restore the original method.
    
    Args:
        verbose: Whether to print information about the operation
        
    Returns:
        bool: True if original method was restored, False otherwise
    """
    global original_method
    
    # Check if the patch was applied
    if original_method is None:
        if verbose:
            print("Patch is not applied, nothing to remove.")
        return True
    
    # Get the module and class
    module, class_object = get_module_and_class()
    if module is None or class_object is None:
        return False
    
    try:
        # Restore the original method
        setattr(class_object, METHOD_NAME, original_method)
        
        # Reset the original method reference
        original_method = None
        
        if verbose:
            print(f"Successfully removed patch from {MODULE_PATH}.{CLASS_NAME}.{METHOD_NAME}")
        
        return True
    except Exception as e:
        print(f"Error removing patch: {e}")
        return False

def is_patch_applied():
    """
    Check if the patch is currently applied.
    
    Returns:
        bool: True if the patch is applied, False otherwise
    """
    return original_method is not None

def example_usage():
    """
    Example showing how to use the patch.
    """
    from docx import Document
    
    # Apply the patch
    if apply_patch():
        print("Patch applied successfully")
    else:
        print("Failed to apply patch")
        return
    
    try:
        # Create a test document
        doc = Document()
        table = doc.add_table(rows=4, cols=4)
        
        # This would normally raise "requested span not rectangular"
        # but now it will work with our patch
        print("Merging cells...")
        
        # Merge a complex shape
        # First merge: top-left 2x2 block
        table.cell(0, 0).merge(table.cell(1, 1))
        
        # Second merge: top-right block that would normally cause an error
        table.cell(0, 2).merge(table.cell(2, 3))
        
        # Save the document
        doc.save('patched_merge_example.docx')
        print("Document saved as 'patched_merge_example.docx'")
        
    finally:
        # Remove the patch when done
        if remove_patch():
            print("Patch removed successfully")
        else:
            print("Failed to remove patch")

if __name__ == "__main__":
    example_usage()
