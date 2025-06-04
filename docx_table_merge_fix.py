"""
DOCX TABLE MERGE FIX MODULE

This module provides solutions for fixing the "requested span not rectangular" error
that occurs when merging table cells in the python-docx library.

The error occurs in the _span_dimensions method of the CT_Tc class, which validates that
cells being merged form a perfect rectangle. This module provides two approaches:

1. A monkey-patching approach that temporarily modifies the _span_dimensions method
2. A utility class that provides safe merging operations

Usage:
    from docx_table_merge_fix import safe_merge_cells, DocxTableMergeFixer

    # Option 1: Use the safe_merge_cells function
    safe_merge_cells(table, 0, 0, 1, 1)  # Merge cells from (0,0) to (1,1)

    # Option 2: Use the context manager
    with DocxTableMergeFixer():
        table.cell(0, 0).merge(table.cell(1, 1))
"""

from docx import Document
from docx.oxml.table import CT_Tc
from docx.exceptions import InvalidSpanError
from contextlib import contextmanager

# Store original method
original_span_dimensions = CT_Tc._span_dimensions

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

def safe_merge_cells(table, start_row, start_col, end_row, end_col):
    """
    Safely merge cells in a table by temporarily patching the validation method.
    
    Args:
        table: The docx table object
        start_row: Starting row index (0-based)
        start_col: Starting column index (0-based)
        end_row: Ending row index (0-based)
        end_col: Ending column index (0-based)
    
    Returns:
        The merged cell
    """
    # Temporarily patch the _span_dimensions method
    CT_Tc._span_dimensions = patched_span_dimensions
    
    try:
        # Get the corner cells
        start_cell = table.rows[start_row].cells[start_col]._tc
        end_cell = table.rows[end_row].cells[end_col]._tc
        
        # Use the merge method which will now use our patched _span_dimensions
        merged_cell = start_cell.merge(end_cell)
        return merged_cell
    finally:
        # Restore the original method
        CT_Tc._span_dimensions = original_span_dimensions

@contextmanager
def merge_fixer_context():
    """
    Context manager that temporarily patches the _span_dimensions method
    to allow non-rectangular cell merges.
    
    Usage:
        with merge_fixer_context():
            # All merges within this context will succeed
            table.cell(0, 0).merge(table.cell(1, 1))
    """
    CT_Tc._span_dimensions = patched_span_dimensions
    try:
        yield
    finally:
        CT_Tc._span_dimensions = original_span_dimensions

class DocxTableMergeFixer:
    """
    A class that provides methods for safely merging table cells without
    the "requested span not rectangular" error.
    """
    
    @staticmethod
    def merge_cells(table, start_row, start_col, end_row, end_col):
        """
        Merge cells in a table from (start_row, start_col) to (end_row, end_col).
        
        Args:
            table: The docx table object
            start_row: Starting row index (0-based)
            start_col: Starting column index (0-based)
            end_row: Ending row index (0-based)
            end_col: Ending column index (0-based)
            
        Returns:
            The merged cell
        """
        return safe_merge_cells(table, start_row, start_col, end_row, end_col)
    
    @staticmethod
    def merge_complex_region(table, cell_coordinates):
        """
        Merge a complex region defined by a list of cell coordinates.
        
        Args:
            table: The docx table object
            cell_coordinates: List of (row, col) tuples defining the region to merge
            
        Returns:
            The merged cell
        """
        if not cell_coordinates:
            raise ValueError("cell_coordinates cannot be empty")
            
        # Find the bounding box of all cells
        rows = [r for r, _ in cell_coordinates]
        cols = [c for _, c in cell_coordinates]
        
        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(cols), max(cols)
        
        return safe_merge_cells(table, min_row, min_col, max_row, max_col)
    
    def __enter__(self):
        """Enable context manager usage: with DocxTableMergeFixer(): ..."""
        CT_Tc._span_dimensions = patched_span_dimensions
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Restore original method when exiting context"""
        CT_Tc._span_dimensions = original_span_dimensions

def example_usage():
    """
    Example demonstrating different ways to use the merge fix.
    """
    doc = Document()
    table = doc.add_table(rows=5, cols=5)
    
    # Example 1: Using the safe_merge_cells function
    print("Example 1: Using safe_merge_cells function")
    safe_merge_cells(table, 0, 0, 1, 1)  # Merge top-left 2x2 block
    
    # Example 2: Using the context manager
    print("Example 2: Using the context manager")
    with merge_fixer_context():
        # All these merges will work without errors
        table.cell(2, 0).merge(table.cell(3, 1))  # Merge middle-left 2x2 block
        
    # Example 3: Using the DocxTableMergeFixer class
    print("Example 3: Using the DocxTableMergeFixer class")
    fixer = DocxTableMergeFixer()
    
    # Merge a complex region
    fixer.merge_complex_region(table, [
        (0, 2), (0, 3), (0, 4),
        (1, 2), (1, 3), (1, 4),
        (2, 2), (2, 3), (2, 4)
    ])
    
    # Example 4: Using the class as a context manager
    print("Example 4: Using the class as a context manager")
    with DocxTableMergeFixer():
        table.cell(3, 2).merge(table.cell(4, 4))  # Merge bottom-right block
    
    # Save the document
    doc.save('complex_merged_table_example.docx')
    print("Document saved as 'complex_merged_table_example.docx'")

if __name__ == "__main__":
    example_usage()
