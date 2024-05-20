import json
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def add_custom_style(doc, style_name, font_name, font_size, bold=False, color=None):
    styles = doc.styles
    if style_name not in styles:
        style = styles.add_style(style_name, 1)  # 1 for paragraph style
        font = style.font
        font.name = font_name
        font.size = Pt(font_size)
        font.bold = bold
        if color:
            font.color.rgb = RGBColor(color[0], color[1], color[2])
    return styles[style_name]

def set_cell_color(cell, color):
    """Set background color of a cell."""
    cell_properties = cell._element.get_or_add_tcPr()
    cell_shading = OxmlElement('w:shd')
    cell_shading.set(qn('w:fill'), color)
    cell_properties.append(cell_shading)

#Get the directory of the Flask application file
app_dir = os.path.dirname(__file__)

# JSON file Path
json_path = os.path.join(app_dir, "result", "json")

# Find the JSON file in the directory
json_file_path = None
for filename in os.listdir(json_path):
    if filename.endswith(".json"):
        json_file_path = os.path.join(json_path, filename)
        break

# Check if a JSON file was found
if json_file_path is None:
    print("No JSON file found in the directory.")
    exit()
# Load JSON data from file
with open(json_file_path, 'r') as file:
    json_data = json.load(file)

# Create a new Document
doc = Document()

# Add heading
heading = doc.add_heading(level=1)
run = heading.add_run('Appendix A.1 Certificate Details')
run.bold = True
run.font.size = Pt(14)
heading.alignment = WD_ALIGN_PARAGRAPH.LEFT


# Add a table
table = doc.add_table(rows=0, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = 'Table Grid'

# Add header row with merged cells for title
header_cells = table.add_row().cells
header_cells[0].merge(header_cells[1])
header_run = header_cells[0].paragraphs[0].add_run('Certificate')
header_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
header_run.bold = True
header_run.font.size = Pt(12)
header_run.font.color.rgb = RGBColor(255, 255, 255)
set_cell_color(header_cells[0], '4F81BD')

# Find the 'ip' value for 'cert_commonName'
common_name_item = next((item for item in json_data if item['id'] == 'cert_commonName'), None)
if common_name_item:
    ip_value = common_name_item['ip']
    port_value = common_name_item['port']
    ip_port_value = f"{ip_value}:{port_value}"
else:
    ip_port_value = 'Not Available'

# Add second header row with the dynamically found IP value
url_cells = table.add_row().cells
url_cells[0].merge(url_cells[1])
url_run = url_cells[0].paragraphs[0].add_run(ip__port_value)
url_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
url_run.bold = True
url_run.font.size = Pt(12)
url_run.font.color.rgb = RGBColor(255, 255, 255)
set_cell_color(url_cells[0], '4F81BD')

# Define the row color mappings for severity
severity_color_map = {
    "OK": "FFFFFF",  # Plain White
    "INFO": "FFFFFF",  # Plain White
    "LOW": "D9EAD3",  # Green
    "MEDIUM": "FFF2CC",  # Yellow
    "WARN": "F4CCCC",  # Yellow
    "CRITICAL": "F4CCCC"  # Red
}

# Define the mappings of the keys to search in the JSON file
key_to_id_map = {
    "Serial Number": "cert_serialNumber",
    "Subject": "cert_commonName",
    "Issuer": "cert_caIssuerssubject",
    "Certificate Chain": "cert_trust",
    "Valid From": "cert_notBefore",
    "Valid Until": "cert_notAfter",
    "Signature Algorithm": "cert_signatureAlgorithm",
    "Public Key Size": "cert_keySize",
    "DNS Certificate Authority Authorisation Record": "DNS_CAArecord",
    "Hostname Validation": "cert_hostNameValidation",
    "Self-Signed": "cert_selfSigned",
    "OCSP URL": "cert_ocspURL",
    "Subject Alternative Name": "cert_subjectAltName",
    "OCSP Stapling": "OCSP_stapling",
    "OCSP Must Staple Extension": "cert_mustStapleExtension"
}

# Add the JSON data to the table
for idx, (key, json_id) in enumerate(key_to_id_map.items()):
    row_cells = table.add_row().cells
    key_cell = row_cells[0]
    value_cell = row_cells[1]

    key_run = key_cell.paragraphs[0].add_run(key)
    key_run.font.size = Pt(10)
    key_run.bold = True
    
    # Find the corresponding value and severity for the given id in the JSON data
    item = next((item for item in json_data if item['id'] == json_id), None)
    if item:
        value = item['finding']
        severity = item['severity']
        # Replacing -- value with N/A
        if value == '--':
            value = 'N/A'
    else:
        value = 'Not Available'
        severity = 'INFO'  # Default to INFO if not found

    value_run = value_cell.paragraphs[0].add_run(value)
    value_run.font.size = Pt(10)

    # Set background color based on severity
    bg_color = severity_color_map.get(severity, "FFFFFF")
    set_cell_color(value_cell, bg_color)

# Save the document
doc.save('Certificate_Details_Report.docx')
