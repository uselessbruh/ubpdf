"""
Universal Document Converter
Handles: PDF ↔ Word, PDF ↔ Excel, HTML → PDF, Office → PDF
Usage: converter.exe <conversion_type> <input_file> <output_file>
"""

import sys
import os
from pathlib import Path

def pdf_to_word(input_path, output_path):
    """Convert PDF to Word DOCX with enhanced formatting preservation"""
    from pdf2docx import Converter
    
    cv = Converter(input_path)
    # Enhanced conversion with better layout, formatting, and table detection
    cv.convert(output_path, 
               start=0, 
               end=None,
               pages=None,
               multi_processing=False,
               cpu_count=1)
    cv.close()
    print(f"✓ Converted PDF to Word: {output_path}")

def word_to_pdf(input_path, output_path):
    """Convert Word DOCX to PDF with formatting preservation"""
    from docx import Document
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, ListFlowable, ListItem
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    
    doc = Document(input_path)
    pdf = SimpleDocTemplate(output_path, pagesize=letter,
                           topMargin=0.75*inch, bottomMargin=0.75*inch,
                           leftMargin=0.75*inch, rightMargin=0.75*inch)
    styles = getSampleStyleSheet()
    story = []
    
    # Add custom styles for different text formats
    styles.add(ParagraphStyle(name='CustomBold', parent=styles['Normal'], 
                              fontName='Helvetica-Bold'))
    styles.add(ParagraphStyle(name='CustomItalic', parent=styles['Normal'], 
                              fontName='Helvetica-Oblique'))
    styles.add(ParagraphStyle(name='CustomBoldItalic', parent=styles['Normal'], 
                              fontName='Helvetica-BoldOblique'))
    
    # Process paragraphs with formatting
    for para in doc.paragraphs:
        if para.text.strip():
            # Determine style based on runs
            style_name = 'Normal'
            font_size = 12
            
            # Check if paragraph has special formatting
            if para.runs:
                first_run = para.runs[0]
                if first_run.bold and first_run.italic:
                    style_name = 'CustomBoldItalic'
                elif first_run.bold:
                    style_name = 'CustomBold'
                elif first_run.italic:
                    style_name = 'CustomItalic'
                
                # Get font size if available
                if first_run.font.size:
                    font_size = first_run.font.size.pt
            
            # Build formatted text with HTML-like tags for styling
            formatted_text = ""
            for run in para.runs:
                text = run.text
                if run.bold:
                    text = f"<b>{text}</b>"
                if run.italic:
                    text = f"<i>{text}</i>"
                if run.underline:
                    text = f"<u>{text}</u>"
                formatted_text += text
            
            # Create paragraph with appropriate style
            current_style = styles[style_name]
            current_style.fontSize = font_size
            current_style.leading = font_size * 1.2
            
            p = Paragraph(formatted_text or para.text, current_style)
            story.append(p)
            story.append(Spacer(1, 0.2*inch))
    
    # Process tables
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text)
            table_data.append(row_data)
        
        if table_data:
            t = Table(table_data)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(t)
            story.append(Spacer(1, 0.3*inch))
    
    pdf.build(story)
    print(f"✓ Converted Word to PDF: {output_path}")

def pdf_to_excel(input_path, output_path):
    """Extract tables from PDF to Excel with formatting"""
    import pdfplumber
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tables"
    
    current_row = 1
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    with pdfplumber.open(input_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            
            if tables:
                for table in tables:
                    # Page label
                    cell = ws.cell(row=current_row, column=1, value=f"Page {page_num}")
                    cell.font = Font(bold=True, size=12, color='366092')
                    current_row += 1
                    
                    # Table header
                    if table:
                        for col_num, cell_value in enumerate(table[0], 1):
                            cell = ws.cell(row=current_row, column=col_num, value=cell_value)
                            cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
                            cell.font = Font(bold=True, color='FFFFFF')
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = border
                        current_row += 1
                        
                        # Table data
                        for row_data in table[1:]:
                            for col_num, cell_value in enumerate(row_data, 1):
                                cell = ws.cell(row=current_row, column=col_num, value=cell_value)
                                cell.border = border
                                cell.alignment = Alignment(horizontal='left', vertical='center')
                            current_row += 1
                    
                    current_row += 1  # Add space between tables
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(output_path)
    print(f"✓ Extracted tables to Excel: {output_path}")

def excel_to_pdf(input_path, output_path):
    """Convert Excel XLSX to PDF with enhanced table formatting"""
    from openpyxl import load_workbook
    from openpyxl.styles import Font as XLFont, Alignment as XLAlignment
    from reportlab.lib.pagesizes import letter, landscape, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    
    wb = load_workbook(input_path)
    
    # Determine page orientation based on number of columns
    first_sheet = wb.worksheets[0]
    max_cols = first_sheet.max_column
    pagesize = landscape(letter) if max_cols > 8 else letter
    
    pdf = SimpleDocTemplate(output_path, pagesize=pagesize,
                           topMargin=0.5*inch, bottomMargin=0.5*inch,
                           leftMargin=0.5*inch, rightMargin=0.5*inch)
    
    styles = getSampleStyleSheet()
    story = []
    
    # Process each worksheet
    for sheet_idx, ws in enumerate(wb.worksheets):
        if sheet_idx > 0:
            story.append(PageBreak())
        
        # Add sheet title
        if len(wb.worksheets) > 1:
            title_style = ParagraphStyle('SheetTitle', parent=styles['Heading1'],
                                        fontSize=14, textColor=colors.HexColor('#366092'),
                                        spaceAfter=12)
            story.append(Paragraph(f"<b>{ws.title}</b>", title_style))
            story.append(Spacer(1, 0.2*inch))
        
        # Get data from Excel
        data = []
        cell_styles = []
        
        for row_idx, row in enumerate(ws.iter_rows(values_only=False), 1):
            row_data = []
            row_styles = []
            for cell in row:
                # Get cell value
                value = cell.value if cell.value is not None else ''
                row_data.append(str(value))
                
                # Store cell formatting info
                row_styles.append({
                    'bold': cell.font.bold if cell.font else False,
                    'italic': cell.font.italic if cell.font else False,
                    'size': cell.font.size if cell.font and cell.font.size else 10,
                    'color': cell.font.color if cell.font else None,
                    'fill': cell.fill.start_color.rgb if cell.fill and cell.fill.start_color else None,
                    'alignment': cell.alignment.horizontal if cell.alignment else 'left'
                })
            
            data.append(row_data)
            cell_styles.append(row_styles)
            
            # Limit rows to prevent huge PDFs
            if row_idx >= 100:
                break
        
        if not data:
            continue
        
        # Create table with dynamic column widths
        available_width = pagesize[0] - inch
        col_widths = [available_width / len(data[0]) for _ in range(len(data[0]))]
        
        t = Table(data, colWidths=col_widths)
        
        # Build table style based on Excel formatting
        table_style_commands = [
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ]
        
        # Apply header formatting (first row)
        if cell_styles and cell_styles[0]:
            header_bold = any(style.get('bold', False) for style in cell_styles[0])
            if header_bold or True:  # Default header formatting
                table_style_commands.extend([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#366092')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ])
        
        # Apply cell-specific formatting
        for row_idx, row_style in enumerate(cell_styles[1:], 1):  # Skip header
            for col_idx, cell_style in enumerate(row_style):
                if cell_style.get('bold'):
                    table_style_commands.append(
                        ('FONTNAME', (col_idx, row_idx), (col_idx, row_idx), 'Helvetica-Bold')
                    )
                if cell_style.get('fill') and cell_style['fill'] != 'FFFFFFFF':
                    # Apply background color if not white
                    try:
                        bg_color = colors.HexColor('#' + cell_style['fill'][-6:])
                        table_style_commands.append(
                            ('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), bg_color)
                        )
                    except:
                        pass
        
        # Alternate row colors for better readability
        for row_idx in range(1, len(data), 2):
            table_style_commands.append(
                ('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor('#F5F5F5'))
            )
        
        t.setStyle(TableStyle(table_style_commands))
        story.append(t)
        story.append(Spacer(1, 0.3*inch))
    
    pdf.build(story)
    print(f"✓ Converted Excel to PDF: {output_path}")

def html_to_pdf(input_path, output_path):
    """Convert HTML to PDF with proper table, header, and font support"""
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    from html.parser import HTMLParser
    from html import unescape
    import re
    
    class EnhancedHTMLParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.elements = []
            self.current_text = []
            self.in_bold = False
            self.in_italic = False
            self.in_heading = 0
            self.in_table = False
            self.in_table_row = False
            self.in_table_header = False
            self.current_table = []
            self.current_row = []
            self.current_cell = []
            self.font_stack = []
            
        def handle_starttag(self, tag, attrs):
            attrs_dict = dict(attrs)
            
            if tag == 'table':
                self.in_table = True
                self.current_table = []
            elif tag == 'tr':
                self.in_table_row = True
                self.current_row = []
            elif tag in ['th', 'td']:
                self.in_table_header = (tag == 'th')
                self.current_cell = []
            elif tag in ['b', 'strong']:
                self.in_bold = True
            elif tag in ['i', 'em']:
                self.in_italic = True
            elif tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                self.in_heading = int(tag[1])
            elif tag == 'p':
                if self.current_text:
                    self.flush_text()
            elif tag == 'br':
                self.current_text.append('<br/>')
                
        def handle_endtag(self, tag):
            if tag == 'table':
                if self.current_table:
                    self.elements.append(('table', self.current_table))
                self.in_table = False
            elif tag == 'tr':
                if self.current_row:
                    self.current_table.append(self.current_row)
                self.in_table_row = False
            elif tag in ['th', 'td']:
                cell_text = ' '.join(self.current_cell).strip()
                self.current_row.append({
                    'text': cell_text,
                    'is_header': self.in_table_header
                })
                self.current_cell = []
                self.in_table_header = False
            elif tag in ['b', 'strong']:
                self.in_bold = False
            elif tag in ['i', 'em']:
                self.in_italic = False
            elif tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                self.flush_text()
                self.in_heading = 0
            elif tag in ['p', 'div']:
                self.flush_text()
                
        def handle_data(self, data):
            data = data.strip()
            if not data:
                return
                
            if self.in_table:
                self.current_cell.append(data)
            else:
                # Format text based on current state
                formatted = data
                if self.in_bold:
                    formatted = f'<b>{formatted}</b>'
                if self.in_italic:
                    formatted = f'<i>{formatted}</i>'
                    
                self.current_text.append(formatted)
                
        def flush_text(self):
            if self.current_text:
                text = ' '.join(self.current_text)
                style_type = f'h{self.in_heading}' if self.in_heading else 'normal'
                self.elements.append((style_type, text))
                self.current_text = []
    
    # Read HTML file
    with open(input_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # Parse HTML
    parser = EnhancedHTMLParser()
    parser.feed(html_content)
    parser.flush_text()
    
    # Create PDF
    pdf = SimpleDocTemplate(output_path, pagesize=letter,
                           topMargin=0.75*inch, bottomMargin=0.75*inch,
                           leftMargin=0.75*inch, rightMargin=0.75*inch)
    
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Heading1Custom', parent=styles['Heading1'],
                              fontSize=18, spaceAfter=12, textColor=colors.HexColor('#366092')))
    styles.add(ParagraphStyle(name='Heading2Custom', parent=styles['Heading2'],
                              fontSize=16, spaceAfter=10, textColor=colors.HexColor('#366092')))
    
    story = []
    
    for element_type, content in parser.elements:
        if element_type == 'table':
            # Build table data
            table_data = []
            for row in content:
                row_data = [cell['text'] for cell in row]
                table_data.append(row_data)
            
            if table_data:
                t = Table(table_data)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#366092')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                story.append(t)
                story.append(Spacer(1, 0.3*inch))
        elif element_type.startswith('h'):
            # Heading
            level = element_type[1]
            style_name = f'Heading{level}Custom' if f'Heading{level}Custom' in styles else 'Heading1'
            p = Paragraph(content, styles.get(style_name, styles['Heading1Custom']))
            story.append(p)
            story.append(Spacer(1, 0.1*inch))
        else:
            # Normal paragraph
            if content.strip():
                p = Paragraph(content, styles['Normal'])
                story.append(p)
                story.append(Spacer(1, 0.1*inch))
    
    pdf.build(story)
    print(f"✓ Converted HTML to PDF: {output_path}")

def ppt_to_pdf(input_path, output_path):
    """Convert PowerPoint to PDF with enhanced formatting"""
    from pptx import Presentation
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    
    prs = Presentation(input_path)
    pdf = SimpleDocTemplate(output_path, pagesize=landscape(letter),
                           topMargin=0.5*inch, bottomMargin=0.5*inch,
                           leftMargin=0.75*inch, rightMargin=0.75*inch)
    styles = getSampleStyleSheet()
    story = []
    
    # Custom styles
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'],
                                 fontSize=24, textColor=colors.HexColor('#366092'),
                                 alignment=TA_CENTER, spaceAfter=20,
                                 fontName='Helvetica-Bold')
    
    heading_style = ParagraphStyle('CustomHeading', parent=styles['Heading2'],
                                   fontSize=18, textColor=colors.HexColor('#366092'),
                                   spaceAfter=12, fontName='Helvetica-Bold')
    
    bullet_style = ParagraphStyle('CustomBullet', parent=styles['Normal'],
                                  fontSize=12, leftIndent=20, spaceAfter=8,
                                  bulletIndent=10)
    
    for slide_num, slide in enumerate(prs.slides, 1):
        if slide_num > 1:
            story.append(PageBreak())
        
        # Add slide number
        slide_label = Paragraph(f"<font size=10 color='gray'>Slide {slide_num}</font>", styles['Normal'])
        story.append(slide_label)
        story.append(Spacer(1, 0.1*inch))
        
        # Process shapes in slide
        for shape in slide.shapes:
            # Handle text frames
            if hasattr(shape, "text") and shape.text.strip():
                text = shape.text.strip()
                
                # Check if it's a title (usually the first text shape)
                if shape == slide.shapes[0] and len(text) < 100:
                    p = Paragraph(f"<b>{text}</b>", title_style)
                    story.append(p)
                    story.append(Spacer(1, 0.2*inch))
                else:
                    # Check if text frame has formatting
                    if hasattr(shape, "text_frame"):
                        text_frame = shape.text_frame
                        
                        for paragraph in text_frame.paragraphs:
                            if paragraph.text.strip():
                                para_text = paragraph.text
                                
                                # Check for bullet points
                                if paragraph.level is not None and paragraph.level > 0:
                                    indent = paragraph.level * 20
                                    custom_bullet = ParagraphStyle('TempBullet', parent=bullet_style,
                                                                  leftIndent=indent,
                                                                  bulletIndent=indent-10)
                                    p = Paragraph(f"• {para_text}", custom_bullet)
                                else:
                                    # Apply run formatting
                                    formatted_text = ""
                                    for run in paragraph.runs:
                                        run_text = run.text
                                        if run.font.bold:
                                            run_text = f"<b>{run_text}</b>"
                                        if run.font.italic:
                                            run_text = f"<i>{run_text}</i>"
                                        if run.font.underline:
                                            run_text = f"<u>{run_text}</u>"
                                        formatted_text += run_text
                                    
                                    p = Paragraph(formatted_text or para_text, styles['Normal'])
                                
                                story.append(p)
                                story.append(Spacer(1, 0.1*inch))
                    else:
                        p = Paragraph(text, styles['Normal'])
                        story.append(p)
                        story.append(Spacer(1, 0.1*inch))
            
            # Handle tables
            elif hasattr(shape, "table"):
                table = shape.table
                table_data = []
                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        row_data.append(cell.text)
                    table_data.append(row_data)
                
                if table_data:
                    t = Table(table_data)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#366092')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                        ('TOPPADDING', (0, 0), (-1, -1), 8),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ]))
                    story.append(t)
                    story.append(Spacer(1, 0.2*inch))
        
        story.append(Spacer(1, 0.3*inch))
    
    pdf.build(story)
    print(f"✓ Converted PowerPoint to PDF: {output_path}")

def pdf_to_html(input_path, output_path):
    """Convert PDF to HTML with better formatting preservation"""
    import pdfplumber
    
    html_content = """<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {
            font-family: Arial, Helvetica, sans-serif;
            padding: 40px;
            line-height: 1.6;
            max-width: 1200px;
            margin: 0 auto;
            background-color: #f5f5f5;
        }
        .page {
            background: white;
            padding: 40px;
            margin-bottom: 30px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            page-break-after: always;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        table th {
            background-color: #366092;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
        }
        table td {
            border: 1px solid #ddd;
            padding: 10px;
        }
        table tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        h1, h2, h3 {
            color: #333;
            margin-top: 20px;
        }
        p {
            margin: 10px 0;
        }
        .bold { font-weight: bold; }
        .italic { font-style: italic; }
        .underline { text-decoration: underline; }
        ul, ol {
            margin: 10px 0;
            padding-left: 30px;
        }
        li {
            margin: 5px 0;
        }
    </style>
</head>
<body>
"""
    
    with pdfplumber.open(input_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            html_content += f"<div class='page'>"
            html_content += f"<p style='color: #999; font-size: 12px;'>Page {page_num}</p>"
            
            # Extract text with layout
            text = page.extract_text(layout=True)
            if text:
                # Convert text to HTML paragraphs
                paragraphs = text.split('\n\n')
                for para in paragraphs:
                    if para.strip():
                        # Detect lists
                        lines = para.split('\n')
                        if any(line.strip().startswith(('•', '-', '*', '●', '○')) for line in lines):
                            html_content += "<ul>"
                            for line in lines:
                                if line.strip():
                                    clean_line = line.strip().lstrip('•-*●○ ')
                                    html_content += f"<li>{clean_line}</li>"
                            html_content += "</ul>"
                        elif any(line.strip() and line.strip()[0].isdigit() and '.' in line[:5] for line in lines):
                            html_content += "<ol>"
                            for line in lines:
                                if line.strip():
                                    # Remove number prefix
                                    clean_line = line.strip()
                                    if clean_line and clean_line[0].isdigit():
                                        clean_line = '.'.join(clean_line.split('.')[1:]).strip()
                                    html_content += f"<li>{clean_line}</li>"
                            html_content += "</ol>"
                        else:
                            # Regular paragraph
                            html_content += f"<p>{para.replace(chr(10), '<br>')}</p>"
            
            # Extract tables
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    html_content += "<table>"
                    for i, row in enumerate(table):
                        if i == 0:
                            # Header row
                            html_content += "<tr>"
                            for cell in row:
                                html_content += f"<th>{cell if cell else ''}</th>"
                            html_content += "</tr>"
                        else:
                            # Data rows
                            html_content += "<tr>"
                            for cell in row:
                                html_content += f"<td>{cell if cell else ''}</td>"
                            html_content += "</tr>"
                    html_content += "</table>"
            
            html_content += "</div>"
    
    html_content += "</body></html>"
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✓ Converted PDF to HTML: {output_path}")

def pdf_to_ppt(input_path, output_path):
    """Convert PDF to PowerPoint with better text extraction and formatting"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    import pdfplumber
    
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    with pdfplumber.open(input_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # Use blank layout
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
            
            # Add page number in corner
            page_num_box = slide.shapes.add_textbox(
                Inches(9), Inches(7), Inches(0.8), Inches(0.3)
            )
            page_num_frame = page_num_box.text_frame
            page_num_frame.text = f"Page {page_num}"
            page_num_para = page_num_frame.paragraphs[0]
            page_num_para.font.size = Pt(10)
            page_num_para.font.color.rgb = RGBColor(128, 128, 128)
            
            # Extract text with layout preservation
            text = page.extract_text(layout=True)
            
            if text:
                # Create main text box
                textbox = slide.shapes.add_textbox(
                    Inches(0.5), Inches(0.5), Inches(9), Inches(6.5)
                )
                text_frame = textbox.text_frame
                text_frame.word_wrap = True
                
                # Split into lines and paragraphs
                lines = text.split('\n')
                first_para = True
                
                for line in lines:
                    if line.strip():
                        if first_para:
                            text_frame.text = line
                            p = text_frame.paragraphs[0]
                            first_para = False
                        else:
                            p = text_frame.add_paragraph()
                            p.text = line
                        
                        # Detect and format titles (short, uppercase, or first line)
                        if len(line.strip()) < 60 and (line.isupper() or lines.index(line) == 0):
                            p.font.size = Pt(24)
                            p.font.bold = True
                            p.font.color.rgb = RGBColor(54, 96, 146)
                            p.alignment = PP_ALIGN.CENTER
                        # Detect bullets
                        elif line.strip().startswith(('•', '-', '*', '●', '○')):
                            p.text = line.strip().lstrip('•-*●○ ')
                            p.level = 1
                            p.font.size = Pt(14)
                        # Regular text
                        else:
                            p.font.size = Pt(12)
                        
                        # Add spacing
                        p.space_after = Pt(6)
            
            # Extract and add tables
            tables = page.extract_tables()
            if tables and len(tables) > 0:
                # Add table to slide (simplified - put in separate slide if text exists)
                table = tables[0]
                if len(table) > 0 and len(table[0]) > 0:
                    # Create new slide for table
                    table_slide = prs.slides.add_slide(prs.slide_layouts[6])
                    
                    # Calculate dimensions
                    rows = min(len(table), 20)
                    cols = min(len(table[0]), 10)
                    
                    # Add table shape
                    left = Inches(0.5)
                    top = Inches(1)
                    width = Inches(9)
                    height = Inches(5)
                    
                    shape = table_slide.shapes.add_table(rows, cols, left, top, width, height)
                    ppt_table = shape.table
                    
                    # Fill table data
                    for i, row in enumerate(table[:rows]):
                        for j, cell_value in enumerate(row[:cols]):
                            cell = ppt_table.cell(i, j)
                            cell.text = str(cell_value) if cell_value else ''
                            
                            # Format header row
                            if i == 0:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = RGBColor(54, 96, 146)
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        run.font.color.rgb = RGBColor(255, 255, 255)
                                        run.font.bold = True
                                        run.font.size = Pt(11)
                            else:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        run.font.size = Pt(10)
    
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
