# Python-DOCX Table Merge Fix

## Problem: "Requested span not rectangular" Error

When working with the `python-docx` library to merge table cells, you may encounter this error:

```
docx.exceptions.InvalidSpanError: requested span not rectangular
```

This error occurs in the `_span_dimensions` method of the `CT_Tc` class when trying to merge cells that don't form a perfect rectangle. The library has strict validation to prevent non-rectangular cell merges.

## Root Cause Analysis

The error occurs in `_span_dimensions` method in `docx/oxml/table.py` which has two validation functions:

1. `raise_on_inverted_L`: Prevents L-shaped merges
2. `raise_on_tee_shaped`: Prevents T-shaped merges

These functions check if the cells you're trying to merge form a perfect rectangle. If not, they raise `InvalidSpanError` with the message "requested span not rectangular".

## Solutions

This repository provides three different solutions to fix this issue:

### 1. Temporary Fix Using Context Manager (docx_table_merge_fix.py)

```python
from docx_table_merge_fix import DocxTableMergeFixer, safe_merge_cells

# Option 1: Use the safe_merge_cells function
safe_merge_cells(table, 0, 0, 1, 1)  # Merge cells from (0,0) to (1,1)

# Option 2: Use the context manager
with DocxTableMergeFixer():
    table.cell(0, 0).merge(table.cell(1, 1))  # This won't raise an error
```

### 2. Permanent Library Patch (docx_table_merge_patch.py)

```python
from docx_table_merge_patch import apply_patch, remove_patch

# Apply the patch at the start of your application
apply_patch()

# Your code using python-docx here...
# All table merges will work without the rectangular error

# Optionally, remove the patch when done
remove_patch()
```

### 3. Integration with PDF Converter (integrate_docx_merge_fix.py)

See the `integrate_docx_merge_fix.py` file for an example of how to integrate the fix with your PDF converter application.

## How the Fix Works

The fix works by replacing or temporarily modifying the `_span_dimensions` method in the `CT_Tc` class to bypass the rectangular validation. Instead of checking for perfect rectangles, it:

1. Calculates the minimum bounding rectangle that encompasses all selected cells
2. Returns those dimensions for the merge operation

This allows merging cells that don't form a perfect rectangle, while still maintaining the table structure.

## Installation

No installation is required. Simply copy the necessary files to your project:

- `docx_table_merge_fix.py` - For the context manager approach
- `docx_table_merge_patch.py` - For the permanent patch approach
- `integrate_docx_merge_fix.py` - For integration examples

## Usage Examples

### Basic Usage

```python
from docx import Document
from docx_table_merge_fix import safe_merge_cells

# Create a document with a table
doc = Document()
table = doc.add_table(rows=4, cols=4)

# Merge cells safely
safe_merge_cells(table, 0, 0, 1, 1)  # Top-left 2x2 block
safe_merge_cells(table, 0, 2, 2, 3)  # Right side block that would normally cause an error

# Save the document
doc.save('merged_table_example.docx')
```

### Using the Context Manager

```python
from docx import Document
from docx_table_merge_fix import DocxTableMergeFixer

# Create a document with a table
doc = Document()
table = doc.add_table(rows=4, cols=4)

# Use the context manager for multiple operations
with DocxTableMergeFixer():
    # All these merges will work without errors
    table.cell(0, 0).merge(table.cell(1, 1))  # Top-left 2x2 block
    table.cell(0, 2).merge(table.cell(2, 3))  # Right side block 
    table.cell(3, 0).merge(table.cell(3, 3))  # Bottom row

# Save the document
doc.save('merged_table_example.docx')
```

### Using the Permanent Patch

```python
from docx import Document
from docx_table_merge_patch import apply_patch, remove_patch

# Apply the patch
apply_patch()

# Create a document with a table
doc = Document()
table = doc.add_table(rows=4, cols=4)

# These merges will all work without errors
table.cell(0, 0).merge(table.cell(1, 1))  # Top-left 2x2 block
table.cell(0, 2).merge(table.cell(2, 3))  # Right side block

# Save the document
doc.save('merged_table_example.docx')

# Remove the patch when done
remove_patch()
```

## Integration with PDF Converter

The `integrate_docx_merge_fix.py` file provides an example of how to integrate this fix with your PDF converter application. It demonstrates:

1. How to import and use the fix in your workflow
2. How to process all tables in a document with the fix applied
3. How to handle potential merge regions based on your table analysis

## Limitations

- The fix bypasses validation that was put in place for a reason. In some cases, non-rectangular merges might cause issues with document rendering in some applications.
- The patch should be used carefully, as it modifies a core function of the library.
- When using the context manager approach, ensure that all merge operations happen within the context.

## License

This fix is provided as-is, without any warranty. Use at your own risk.

## Contribution

Feel free to contribute improvements to this fix or report any issues you encounter.

## Related Issues

This fix addresses a common issue with the python-docx library. For more information, see:
- [python-docx Issue #259](https://github.com/python-openxml/python-docx/issues/259)
- [python-docx Issue #713](https://github.com/python-openxml/python-docx/issues/713)
