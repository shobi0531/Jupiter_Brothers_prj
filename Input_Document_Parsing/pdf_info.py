import sys
import os
import json
import fitz 
import camelot
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

def load_per_rate_json(json_file_path):
    with open(json_file_path, "r") as file:
        data = json.load(file)
    return data["per_rate"]

def extract_pdf_info(pdf_path, per_rate_data):
    doc = fitz.open(pdf_path)

    content = {
        "style": None,
        "style_number": None,
        "brand": None,
        "sizes": None,
        "commodity": None,
        "email": None,
        "care_address": None,
        "main_image_path": None,
        "spec_sheet": None
    }

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        text = page.get_text("text")

        if not content["style"]:
            content["style"] = search_for_property(text, "Style")
        if not content["style_number"]:
            content["style_number"] = search_for_property(text, "Style number")
        if not content["brand"]:
            content["brand"] = search_for_property(text, "Brand")
        if not content["sizes"]:
            content["sizes"] = search_for_property(text, "Sizes")
        if not content["commodity"]:
            content["commodity"] = search_for_property(text, "Commodity")
        if not content["email"]:
            content["email"] = search_for_property(text, "E-mail")
        if not content["care_address"]:
            content["care_address"] = search_for_property(text, "Care Address")
        if not content["main_image_path"]:
            content["main_image_path"] = extract_main_image(page)

    content["spec_sheet"] = extract_spec_sheet_table_with_camelot(pdf_path, per_rate_data)
    return content

def search_for_property(text, property_name):
    lines = text.split("\n")
    for line in lines:
        if property_name in line:
            return line.strip()
    return None

def extract_main_image(page):
    image_list = page.get_images(full=True)
    if image_list:
        xref = image_list[0][0]
        image = fitz.Pixmap(page.parent, xref)
        image_filename = f"extracted_image_{xref}.png"
        image.save(image_filename)
        return image_filename
    return None

def extract_spec_sheet_table_with_camelot(pdf_path, per_rate_data):
    tables = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")

    if len(tables) == 0:
        tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream", split_text=True)

    valid_tables = [table.df for table in tables if table.df.shape[0] > 1 and table.df.shape[1] > 1]

    if not valid_tables:
        return None

    last_table = valid_tables[-1]
    last_table.columns = last_table.iloc[0]
    last_table = last_table[1:]
    last_table.reset_index(drop=True, inplace=True)

    last_table["Per Rate"] = per_rate_data[:len(last_table)]
    last_table["Total"] = last_table["Qty"].astype(float) * last_table["Per Rate"].astype(float)

    columns = list(last_table.columns)
    if len(columns) >= 3:
        columns[1], columns[2] = columns[2], columns[1]
        last_table = last_table[columns]

    total_sum = last_table["Total"].sum()
    total_row = ["", "", "Total", "", total_sum]
    last_table.loc[len(last_table)] = total_row

    return last_table

def wrap_text(canvas, text, x, y, max_width, line_height):
    words = text.split(' ')
    current_line = ""
    lines = []

    for word in words:
        test_line = current_line + word + " "
        if canvas.stringWidth(test_line, "Helvetica", 12) <= max_width:
            current_line = test_line
        else:
            lines.append(current_line.strip())
            current_line = word + " "
    if current_line:
        lines.append(current_line.strip())

    for line in lines:
        canvas.drawString(x, y, line)
        y -= line_height
    return y

def create_pdf(output_path, extracted_data):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    x_margin = 72
    font_size = 12
    line_spacing = font_size * 2.5
    y_position = height - 72

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 40, "Costing Sheet")
    c.setFont("Helvetica", font_size)

    c.drawString(72, y_position, f"Style: {extracted_data['style']}")
    y_position -= line_spacing
    c.drawString(72, y_position, f"Style Number: {extracted_data['style_number']}")
    y_position -= line_spacing
    c.drawString(72, y_position, f"Brand: {extracted_data['brand']}")
    y_position -= line_spacing
    c.drawString(72, y_position, f"Sizes: {extracted_data['sizes']}")
    y_position -= line_spacing
    c.drawString(72, y_position, f"Commodity: {extracted_data['commodity']}")
    y_position -= line_spacing
    c.drawString(72, y_position, f"Email: {extracted_data['email']}")
    y_position -= line_spacing
    y_position = wrap_text(c, f"Care Address: {extracted_data['care_address']}", 72, y_position, width - 144, line_spacing)

    if extracted_data["main_image_path"]:
        try:
            c.drawImage(extracted_data["main_image_path"], 400, y_position + 100, width=150, height=150)
            y_position -= 180
        except Exception as e:
            c.drawString(72, y_position, f"Error loading image: {e}")
            y_position -= 20

    if extracted_data["spec_sheet"] is not None:
        table = extracted_data["spec_sheet"]
        data = [list(table.columns)] + table.values.tolist()

        styles = getSampleStyleSheet()
        for i, row in enumerate(data):
            for j, cell in enumerate(row):
                data[i][j] = Paragraph(str(cell), styles["BodyText"])

        col_widths = [100, 150, 50, 50, 50]
        reportlab_table = Table(data, colWidths=col_widths)
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ])
        reportlab_table.setStyle(style)

        y_table_position = y_position - 100
        reportlab_table.wrapOn(c, width - 2 * x_margin, y_table_position)
        reportlab_table.drawOn(c, x_margin, y_table_position)

    c.save()

def main(input_pdf_path):
    # Ensure the attachments folder exists
    attachments_dir = "attachments"
    if not os.path.exists(attachments_dir):
        try:
            os.makedirs(attachments_dir)
        except Exception as e:
            print(f"Error creating directory '{attachments_dir}': {e}")
            return None

    json_file_path = "per_rate.json"
    per_rate_data = load_per_rate_json(json_file_path)

    if not os.path.exists(input_pdf_path):
        print(f"Input file {input_pdf_path} not found.")
        return None

    try:
        extracted_data = extract_pdf_info(input_pdf_path, per_rate_data)
        output_pdf_path = os.path.join(attachments_dir, "new_10.pdf")
        create_pdf(output_pdf_path, extracted_data)
        return output_pdf_path
    except Exception as e:
        print(f"Error processing the PDF: {e}")
        return None


if __name__ == "__main__":
    input_pdf_path = sys.argv[1]
    output_pdf = main(input_pdf_path)
    if output_pdf:
        print(output_pdf)
