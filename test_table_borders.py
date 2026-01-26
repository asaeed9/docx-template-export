#!/usr/bin/env python3
"""
Test script to verify table borders are explicitly defined in XML.

This script creates a simple table and checks that:
1. Table style is applied (preferably "TableGrid")
2. All borders are explicitly defined in XML (top, bottom, left, right, insideH, insideV)
3. Border properties are deterministic (single, size 4, auto color)
"""

from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx_template_export.models.export_models import WordExportRequest
from docx_template_export.services.word_export_service import export_to_word


def test_table_borders_in_xml():
    """Test that tables have explicit borders in XML."""
    print("=" * 80)
    print("Testing Table Borders in XML")
    print("=" * 80)
    
    # Create a simple template with a placeholder
    template_path = Path("test_table_template.docx")
    output_path = Path("test_table_output.docx")
    
    # Create template
    template_doc = Document()
    template_doc.add_paragraph("{{table_content}}")
    template_doc.save(str(template_path))
    print(f"? Created template: {template_path}")
    
    # Create export request with table
    request = WordExportRequest(
        scalar_fields={},
        block_fields={
            "table_content": """
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
"""
        }
    )
    
    # Export
    result = export_to_word(
        template_path=template_path,
        request=request,
        markdown_mode=True,
        output_path=output_path,
    )
    print(f"? Exported to: {output_path}")
    
    # Load and inspect the output
    doc = Document(str(output_path))
    tables = doc.tables
    
    if not tables:
        print("? ERROR: No tables found in output document")
        return False
    
    print(f"\n? Found {len(tables)} table(s)")
    
    all_passed = True
    
    for idx, table in enumerate(tables, 1):
        print(f"\n--- Table {idx} ---")
        
        # Check table style
        style_name = table.style.name if hasattr(table.style, 'name') else str(table.style)
        print(f"Style: {style_name}")
        
        # Check XML for borders
        tbl_elem = table._tbl
        tbl_pr = None
        
        for child in tbl_elem:
            if child.tag == qn('w:tblPr'):
                tbl_pr = child
                break
        
        if tbl_pr is None:
            print("? ERROR: No w:tblPr element found")
            all_passed = False
            continue
        
        # Check for tblStyle
        tbl_style = None
        for child in tbl_pr:
            if child.tag == qn('w:tblStyle'):
                tbl_style = child
                break
        
        if tbl_style is not None:
            style_val = tbl_style.get(qn('w:val'))
            print(f"? Table style in XML: {style_val}")
        else:
            print("? WARNING: No w:tblStyle element found (style may be applied differently)")
        
        # Check for tblBorders
        tbl_borders = None
        for child in tbl_pr:
            if child.tag == qn('w:tblBorders'):
                tbl_borders = child
                break
        
        if tbl_borders is None:
            print("? ERROR: No w:tblBorders element found")
            all_passed = False
            continue
        
        print(f"? w:tblBorders element found")
        
        # Check all required borders
        required_borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
        found_borders = []
        
        for border_elem in tbl_borders:
            border_name = border_elem.tag.split('}')[-1] if '}' in border_elem.tag else border_elem.tag
            found_borders.append(border_name)
            
            # Check border properties
            val = border_elem.get(qn('w:val'))
            sz = border_elem.get(qn('w:sz'))
            color = border_elem.get(qn('w:color'))
            
            print(f"  - {border_name}: val={val}, sz={sz}, color={color}")
        
        # Verify all required borders are present
        missing_borders = []
        for border_name in required_borders:
            border_qn = qn(f'w:{border_name}')
            found = False
            for border_elem in tbl_borders:
                if border_elem.tag == border_qn:
                    found = True
                    break
            if not found:
                missing_borders.append(border_name)
        
        if missing_borders:
            print(f"? ERROR: Missing borders: {missing_borders}")
            all_passed = False
        else:
            print(f"? All required borders present: {', '.join(required_borders)}")
        
        # Verify border properties are deterministic
        for border_name in required_borders:
            border_qn = qn(f'w:{border_name}')
            for border_elem in tbl_borders:
                if border_elem.tag == border_qn:
                    val = border_elem.get(qn('w:val'))
                    sz = border_elem.get(qn('w:sz'))
                    if val != 'single':
                        print(f"? ERROR: {border_name} val is '{val}', expected 'single'")
                        all_passed = False
                    if sz != '4':
                        print(f"? ERROR: {border_name} sz is '{sz}', expected '4'")
                        all_passed = False
                    break
    
    # Cleanup
    if template_path.exists():
        template_path.unlink()
        print(f"\n? Cleaned up template: {template_path}")
    
    print("\n" + "=" * 80)
    if all_passed:
        print("? ALL TESTS PASSED: Tables have explicit borders in XML")
    else:
        print("? SOME TESTS FAILED: Check output above")
    print("=" * 80)
    
    return all_passed


if __name__ == "__main__":
    success = test_table_borders_in_xml()
    exit(0 if success else 1)
