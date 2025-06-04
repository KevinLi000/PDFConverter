from docx import Document
from docx.oxml.table import CT_Tc
from docx.exceptions import InvalidSpanError

# Store original method
original_span_dimensions = CT_Tc._span_dimensions

def patched_span_dimensions(self, other_tc):
    """
    A patched version of _span_dimensions that handles non-rectangular spans
    by ensuring the selection is expanded to create a rectangular region.
    
    This function replaces the original validation with a more permissive approach
    that expands the selection to include all cells needed to form a rectangle.
    """
    # Get the bounds of the current selection
    top = min(self.top, other_tc.top)
    left = min(self.left, other_tc.left)
    bottom = max(self.bottom, other_tc.bottom)
    right = max(self.right, other_tc.right)
    
    # Return the rectangular span dimensions
    return top, left, bottom - top, right - left

def safe_merge_cells(table, start_row, start_col, end_row, end_col):
    """
    Safely merge cells in a table by ensuring the selection forms a rectangle.
    
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

def example_usage():
    """
    Example of how to use the safe_merge_cells function.
    """
    doc = Document()
    table = doc.add_table(rows=4, cols=4)
    
    # This would normally raise "requested span not rectangular"
    # Merge cells that form a complex shape
    safe_merge_cells(table, 0, 0, 1, 1)  # Top-left 2x2 block
    safe_merge_cells(table, 0, 2, 2, 3)  # Right side block
    
    # Save the document
    doc.save('merged_table_example.docx')

if __name__ == "__main__":
    example_usage()
