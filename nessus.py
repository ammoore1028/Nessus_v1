import csv
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import defaultdict
import matplotlib.pyplot as plt
import os

# Setup argparse to take CSV file as a command-line argument
import argparse
parser = argparse.ArgumentParser(description='Generate vulnerability report from CSV.')
parser.add_argument('csv_file', type=str, help='Path to the CSV file containing vulnerability data')
args = parser.parse_args()

# Extract file name without extension and set output Word document name
output_file = args.csv_file.replace('.csv', '.docx')

# Create a new Word Document
doc = Document()

# Function to get the color for a specific risk level
def get_risk_color(risk):
    colors = {
        'Critical': RGBColor(139, 0, 0),  # Burgundy (عنابي)
        'High': RGBColor(255, 0, 0),      # Red (أحمر)
        'Medium': RGBColor(255, 165, 0),  # Orange (برتقالي)
        'Low': RGBColor(0, 100, 0)        # Dark Green (أخضر غامق)
    }
    return colors.get(risk, RGBColor(0, 0, 0))  # Default is black

# Function to add a page break
def add_page_break():
    doc.add_page_break()

# Function to add vulnerability data to the document
def add_vulnerability_to_doc(vuln_data, vuln_number, affected_hosts):
    vuln_title = doc.add_paragraph()
    vuln_run = vuln_title.add_run(f"{vuln_number}. {vuln_data['Name']}")
    vuln_run.bold = True
    vuln_run.font.size = Pt(14)
    vuln_run.font.color.rgb = RGBColor(0, 0, 255)  # Blue color for the name
    vuln_title.alignment = WD_ALIGN_PARAGRAPH.LEFT

    table = doc.add_table(rows=0, cols=2)  # Start with 0 rows
    table.style = 'Table Grid'
    
    table.columns[0].width = Inches(1.5)
    table.columns[1].width = Inches(4.5)
    
    def add_row_if_data(label, data, risk=None):
        if data and data.strip():  # Only add row if data exists
            row_cells = table.add_row().cells
            row_cells[0].text = label
            row_cells[0].paragraphs[0].runs[0].bold = True
            
            data_run = row_cells[1].paragraphs[0].add_run(data)
            data_run.bold = True  # Make the text bold for risk
            if risk:
                data_run.font.color.rgb = get_risk_color(risk)  # Apply color based on risk level
            set_text_direction_ltr(row_cells[1])

    def set_text_direction_ltr(cell):
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '0')  # Set text direction to LTR
        tcPr.append(bidi)

    risk = vuln_data['Risk']
    add_row_if_data("Risk", vuln_data['Risk'] if vuln_data['Risk'] else 'n/a', risk)
    
    affected_hosts = list(set(affected_hosts))  # Remove duplicates
    affected_hosts_text = ', '.join([f"{host}:{port}" for host, port in affected_hosts])
    add_row_if_data("Affected Hosts", affected_hosts_text if affected_hosts_text else 'n/a')
    
    add_row_if_data("References", vuln_data['CVE'] if vuln_data['CVE'] else 'n/a')

    if vuln_data['Description'] and vuln_data['Description'].strip():
        description_row = table.add_row().cells
        description_row[0].merge(description_row[1])
        description_row[0].text = "Description"
        description_row[0].paragraphs[0].runs[0].bold = True
        description_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        desc_content_row = table.add_row().cells
        desc_content_row[0].merge(desc_content_row[1])
        desc_content_row[0].text = vuln_data['Description']
        desc_content_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        desc_content_row[0].paragraphs[0].paragraph_format.space_before = Pt(10)
        desc_content_row[0].paragraphs[0].paragraph_format.space_after = Pt(10)

    if vuln_data['Solution'] and vuln_data['Solution'].strip():
        recommendation_row = table.add_row().cells
        recommendation_row[0].merge(recommendation_row[1])
        recommendation_row[0].text = "Recommendations"
        recommendation_row[0].paragraphs[0].runs[0].bold = True
        recommendation_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        recommendation_content_row = table.add_row().cells
        recommendation_content_row[0].merge(recommendation_content_row[1])
        recommendation_content_row[0].text = vuln_data['Solution']
        recommendation_content_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        recommendation_content_row[0].paragraphs[0].paragraph_format.space_before = Pt(10)
        recommendation_content_row[0].paragraphs[0].paragraph_format.space_after = Pt(10)

vulnerabilities = defaultdict(list)
risk_count = defaultdict(int)
risk_unique_count = defaultdict(set)
with open(args.csv_file, newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['Risk'] and row['Risk'].strip().lower() != 'none':
            vulnerabilities[row['Name']].append((row['Host'], row['Port'], row['Risk'], row))
            risk_count[row['Risk']] += 1
            risk_unique_count[row['Risk']].add(row['Name'])

risk_order = ['Critical', 'High', 'Medium', 'Low']
sorted_vulnerabilities = sorted(vulnerabilities.items(), key=lambda x: risk_order.index(x[1][0][2]))

# Define the labels and sizes for the unique vulnerabilities
labels = ['Critical', 'High', 'Medium', 'Low']
sizes = [len(risk_unique_count[label]) for label in labels]

# Define custom colors with dark green for "Low"
colors = ['#8B0000', '#FF0000', '#FFA500', '#006400']  # Critical = Burgundy, High = Red, Medium = Orange, Low = Dark Green

# Create pie chart with custom colors
plt.pie(sizes, labels=None, colors=colors, startangle=90, autopct=lambda p: f'{int(p * sum(sizes) / 100)}')
plt.title('Vulnerability Assessment')
plt.legend(labels, loc="best")  # Only show labels without numbers
plt.axis('equal')

image_path = "vulnerability_pie_chart.png"
plt.savefig(image_path)
plt.close()

# Insert pie chart on the first page
doc.add_picture(image_path, width=Inches(5))

# Insert a page break after the pie chart
add_page_break()

vuln_number = 1
for vuln_name, hosts_ports_data in sorted_vulnerabilities:
    first_row = hosts_ports_data[0][3]
    affected_hosts = {(host, port) for host, port, _, _ in hosts_ports_data}
    add_vulnerability_to_doc(first_row, vuln_number, affected_hosts)
    
    # Insert a page break after each vulnerability table
    add_page_break()
    
    vuln_number += 1

doc.save(output_file)
os.remove(image_path)

print(f"Pie chart and vulnerabilities each on separate pages added to {output_file}.")
