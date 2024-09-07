import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.units import inch




def excel_to_pdf(excel_file, pdf_file):
    # Read the Excel file
    df = pd.read_excel(excel_file, engine='openpyxl')
    
    # Convert DataFrame to list of lists
    data = [df.columns.tolist()] + df.values.tolist()
    
    # Create PDF document
    doc = SimpleDocTemplate(pdf_file, pagesize=letter)
    
    # Define table style
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONT', (0, 1), (-1, -1), 'Helvetica'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 1), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
    ])
    
    elements = []
    
    # Handle large tables with pagination
    max_rows_per_page = 25  # Adjust based on content and page size
    num_pages = (len(data) // max_rows_per_page) + (len(data) % max_rows_per_page > 0)
    
    for page_num in range(num_pages):
        start_row = page_num * max_rows_per_page
        end_row = min(start_row + max_rows_per_page, len(data))
        page_data = data[start_row:end_row]
        
        # Create a table for the current page
        page_table = Table(page_data, colWidths=[1.5 * inch] * len(df.columns))
        page_table.setStyle(style)
        
        # Add the table to the elements list
        elements.append(page_table)
        
        # Add page break except for the last page
        if page_num < num_pages - 1:
            elements.append(PageBreak())

    # Build the PDF
    doc.build(elements)

# Example usage
excel_file = r"BlueSquare.xlsx"
pdf_file = r"C:\Users\jainh\Downloads\cloud\Shreystore.pdf"

# Convert the Excel file to PDF
excel_to_pdf(excel_file, pdf_file)
