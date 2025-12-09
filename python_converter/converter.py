"""
Universal Document Converter
Handles: PDF ↔ Word, PDF ↔ Excel, HTML → PDF, Office → PDF
Usage: converter.exe <conversion_type> <input_file> <output_file>
"""

import sys
import os
from pathlib import Path

def pdf_to_word(input_path, output_path):
    """Convert PDF to Word DOCX"""
    from pdf2docx import Converter
    
    cv = Converter(input_path)
    cv.convert(output_path)
    cv.close()
    print(f"✓ Converted PDF to Word: {output_path}")

def word_to_pdf(input_path, output_path):
    """Convert Word DOCX to PDF"""
    from docx import Document
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    
    doc = Document(input_path)
    pdf = SimpleDocTemplate(output_path, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    for para in doc.paragraphs:
        if para.text.strip():
            p = Paragraph(para.text, styles['Normal'])
            story.append(p)
            story.append(Spacer(1, 12))
    
    pdf.build(story)
    print(f"✓ Converted Word to PDF: {output_path}")

def pdf_to_excel(input_path, output_path):
    """Extract tables from PDF to Excel"""
    import camelot
    
    tables = camelot.read_pdf(input_path, pages='all', flavor='lattice')
    
    if len(tables) == 0:
        # Try stream flavor if lattice finds nothing
        tables = camelot.read_pdf(input_path, pages='all', flavor='stream')
    
    if len(tables) > 0:
        # If multiple tables, save to different sheets
        if len(tables) == 1:
            tables[0].to_excel(output_path)
        else:
            with open(output_path.replace('.xlsx', '_tables.xlsx'), 'wb') as f:
                writer = tables[0].to_excel(output_path)
        print(f"✓ Extracted {len(tables)} table(s) to Excel: {output_path}")
    else:
        print("⚠ No tables found in PDF")

def excel_to_pdf(input_path, output_path):
    """Convert Excel XLSX to PDF"""
    from openpyxl import load_workbook
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    from reportlab.lib import colors
    
    wb = load_workbook(input_path)
    ws = wb.active
    
    # Get data from Excel
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))
    
    # Create PDF
    pdf = SimpleDocTemplate(output_path, pagesize=landscape(letter))
    
    # Create table
    t = Table(data)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    pdf.build([t])
    print(f"✓ Converted Excel to PDF: {output_path}")

def html_to_pdf(input_path, output_path):
    """Convert HTML to PDF"""
    from weasyprint import HTML
    
    HTML(filename=input_path).write_pdf(output_path)
    print(f"✓ Converted HTML to PDF: {output_path}")

def ppt_to_pdf(input_path, output_path):
    """Convert PowerPoint to PDF"""
    from pptx import Presentation
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet
    
    prs = Presentation(input_path)
    pdf = SimpleDocTemplate(output_path, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    story = []
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                p = Paragraph(shape.text, styles['Normal'])
                story.append(p)
                story.append(Spacer(1, 12))
        story.append(PageBreak())
    
    pdf.build(story)
    print(f"✓ Converted PowerPoint to PDF: {output_path}")

def pdf_to_html(input_path, output_path):
    """Convert PDF to HTML"""
    import pdfplumber
    
    html_content = "<html><head><meta charset='utf-8'><style>body{font-family:Arial;padding:20px;}</style></head><body>"
    
    with pdfplumber.open(input_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                html_content += f"<div style='page-break-after:always;'><p>{text.replace(chr(10), '<br>')}</p></div>"
    
    html_content += "</body></html>"
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✓ Converted PDF to HTML: {output_path}")

def pdf_to_ppt(input_path, output_path):
    """Convert PDF to PowerPoint (basic text extraction)"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    import pdfplumber
    
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    with pdfplumber.open(input_path) as pdf:
        for page in pdf.pages:
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank layout
            text = page.extract_text()
            
            if text:
                textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(6.5))
                text_frame = textbox.text_frame
                text_frame.text = text
                text_frame.word_wrap = True
    
    prs.save(output_path)
    print(f"✓ Converted PDF to PowerPoint: {output_path}")

# Conversion mapping
CONVERSIONS = {
    'pdf-to-word': pdf_to_word,
    'word-to-pdf': word_to_pdf,
    'pdf-to-excel': pdf_to_excel,
    'excel-to-pdf': excel_to_pdf,
    'html-to-pdf': html_to_pdf,
    'ppt-to-pdf': ppt_to_pdf,
    'pdf-to-html': pdf_to_html,
    'pdf-to-ppt': pdf_to_ppt,
}

def main():
    if len(sys.argv) != 4:
        print("Usage: converter.exe <conversion_type> <input_file> <output_file>")
        print("\nAvailable conversions:")
        for conv_type in CONVERSIONS.keys():
            print(f"  - {conv_type}")
        sys.exit(1)
    
    conversion_type = sys.argv[1].lower()
    input_file = sys.argv[2]
    output_file = sys.argv[3]
    
    # Validate input file exists
    if not os.path.exists(input_file):
        print(f"✗ Error: Input file not found: {input_file}")
        sys.exit(1)
    
    # Get conversion function
    if conversion_type not in CONVERSIONS:
        print(f"✗ Error: Unknown conversion type: {conversion_type}")
        print(f"Available: {', '.join(CONVERSIONS.keys())}")
        sys.exit(1)
    
    try:
        # Run conversion
        converter_func = CONVERSIONS[conversion_type]
        converter_func(input_file, output_file)
        sys.exit(0)
    except Exception as e:
        print(f"✗ Conversion failed: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
