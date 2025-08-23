import csv
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import defaultdict
import matplotlib.pyplot as plt
import csv
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

# Function to set font properties to Aptos
def set_font(run, size=11, bold=False, color=None):
    run.font.name = 'Aptos'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Aptos')  # Ensure compatibility with different locales
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color

# Function to add a page break
def add_page_break():
    doc.add_page_break()

# Function to get the color for a specific impact level
def get_impact_color(impact):
    colors = {
        'Critical': RGBColor(139, 0, 0),  # Burgundy 
        'High': RGBColor(255, 0, 0),      # Red 
        'Medium': RGBColor(255, 165, 0),  # Orange 
        'Low': RGBColor(0, 100, 0)        # Dark Green 
    }
    return colors.get(impact, RGBColor(0, 0, 0))  # Default is black

# Scope
scope_heading = doc.add_paragraph('Scope', style='Heading 2')
scope_text = "The scope of the exercise includes and is limited to the following IP addresses:"
scope_paragraph = doc.add_paragraph(scope_text)
set_font(scope_paragraph.runs[0], size=11)

# Load IP addresses from the CSV file (الصفحة الثانية)
unique_ips = set()

# Read the CSV to extract unique IP addresses from the "Host" column
with open(args.csv_file, newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['IP Address']:
            unique_ips.add(row['IP Address'])

unique_ips = sorted(unique_ips)  # Sort IPs for better readability

# Create a table for IP addresses with 4 columns
num_columns = 4
num_rows = len(unique_ips) // num_columns + (1 if len(unique_ips) % num_columns != 0 else 0)
ip_table = doc.add_table(rows=num_rows, cols=num_columns)
ip_table.style = 'Table Grid'
ip_table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Add IP addresses to the table
row_index = 0
col_index = 0

for ip in unique_ips:
    ip_table.cell(row_index, col_index).text = ip
    set_font(ip_table.cell(row_index, col_index).paragraphs[0].runs[0])
    col_index += 1
    if col_index == num_columns:
        col_index = 0
        row_index += 1


# Load CSV data and sort by impact level (الصفحة الثالثة)
vulnerabilities = defaultdict(list)
impact_count = defaultdict(int)
impact_unique_count = defaultdict(set)
with open(args.csv_file, newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['Severity'] and row['Severity'].strip().lower() != 'none':
            vulnerabilities[row['Plugin Name']].append((row['IP Address'], row['Port'], row['Severity'], row))
            impact_count[row['Severity']] += 1
            impact_unique_count[row['Severity']].add(row['Plugin Name'])

# Sort vulnerabilities based on severity order
impact_order = ['Critical', 'High', 'Medium', 'Low']
sorted_vulnerabilities = sorted(vulnerabilities.items(), key=lambda x: impact_order.index(x[1][0][2]))

# ###############################

# Function to get the color for a specific impact level
def get_impact_color(impact):
    colors = {
        'Critical': RGBColor(139, 0, 0),  # Burgundy 
        'High': RGBColor(255, 0, 0),      # Red 
        'Medium': RGBColor(255, 165, 0),  # Orange 
        'Low': RGBColor(0, 100, 0)        # Dark Green 
    }
    return colors.get(impact, RGBColor(0, 0, 0))  # Default is black

# Read the CSV data into a DataFrame to perform operations on it
df = pd.read_csv(args.csv_file)

# Calculate the number of unique vulnerabilities for each severity level
impact_unique_count = {
    'Critical': df[df['Severity'] == 'Critical']['Plugin Name'].nunique(),
    'High': df[df['Severity'] == 'High']['Plugin Name'].nunique(),
    'Medium': df[df['Severity'] == 'Medium']['Plugin Name'].nunique(),
    'Low': df[df['Severity'] == 'Low']['Plugin Name'].nunique()
}

# Calculate the total number of unique vulnerabilities
total_unique_vulnerabilities = sum(impact_unique_count.values())

# Create pie chart with custom colors
labels = ['Critical', 'High', 'Medium', 'Low']
sizes = [impact_unique_count[label] for label in labels]
colors = ['#8B0000', '#FF0000', '#FFA500', '#006400']  # Critical = Burgundy, High = Red, Medium = Orange, Low = Dark Green

plt.pie(sizes, labels=None, colors=colors, startangle=90, autopct=lambda p: f'{int(p * sum(sizes) / 100)}')
plt.title('Vulnerability Assessment')  # تغيير العنوان إلى "Vulnerability Assessment"
plt.legend(labels, loc="best")  # Only show labels without numbers
plt.axis('equal')

image_path = "impact_pie_chart.png"
plt.savefig(image_path)
plt.close()

# Insert pie chart on the next page
doc.add_picture(image_path, width=Inches(5))

# Insert a page break after the pie chart
add_page_break()

# Continue with the rest of the report generation...
# Here, you'll continue adding other details and sections as per your requirements.

# Save the document
doc.save(output_file)
os.remove(image_path)


# ###############################
# Insert a heading before the summary table (الصفحة الثالثة)
heading = doc.add_paragraph('Table of All Vulnerabilities', style='Heading 1')
set_font(heading.runs[0], size=18, bold=True)

# Add a summary table of vulnerabilities before the pie chart
summary_table = doc.add_table(rows=1, cols=3)
summary_table.style = 'Table Grid'

# Set headers for the summary table with coloring and formatting
hdr_cells = summary_table.rows[0].cells
hdr_cells[0].text = '#'
hdr_cells[1].text = 'Finding'
hdr_cells[2].text = 'Risk'

# Color the first line of the table as specified (using black background and white text)
for i, cell in enumerate(hdr_cells):
    set_font(cell.paragraphs[0].runs[0], bold=True, color=RGBColor(255, 255, 255))  # White text color
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), '000000')  # Black color for header background
    cell._element.get_or_add_tcPr().append(shading_elm)
    if i == 2:  # Center align the Impact column
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# Add the data to the summary table
vuln_number = 1
for vuln_name, hosts_ports_data in sorted_vulnerabilities:
    row_cells = summary_table.add_row().cells
    row_cells[0].text = str(vuln_number)
    set_font(row_cells[0].paragraphs[0].runs[0])

    row_cells[1].text = vuln_name
    set_font(row_cells[1].paragraphs[0].runs[0])

    impact_level = hosts_ports_data[0][2]
    row_cells[2].text = impact_level
    impact_run = row_cells[2].paragraphs[0].runs[0]
    set_font(impact_run, bold=True, color=RGBColor(255, 255, 255))  # White text color for impact level
    row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # Align impact in the center

    # Set background color for the impact cell based on the impact level
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), {
        'Critical': '8B0000',  # Burgundy
        'High': 'FF0000',      # Red
        'Medium': 'FFA500',    # Orange
        'Low': '006400'        # Dark Green
    }[impact_level])
    row_cells[2]._element.get_or_add_tcPr().append(shading_elm)

    vuln_number += 1

# Add a page break after the summary table
add_page_break()




# Add detailed vulnerabilities
vuln_number = 1
for vuln_name, hosts_ports_data in sorted_vulnerabilities:
    first_row = hosts_ports_data[0][3]
    affected_hosts = {(host, port) for host, port, _, _ in hosts_ports_data}
    
    # Add vulnerability to document
    vuln_title = doc.add_paragraph(f"{vuln_number}. {first_row['Plugin Name']}", style='Heading 2')
    set_font(vuln_title.runs[0], size=14, bold=True, color=RGBColor(0, 0, 255))  # Blue color for the name
    vuln_title.alignment = WD_ALIGN_PARAGRAPH.LEFT

    table = doc.add_table(rows=0, cols=2)  # Start with 0 rows
    table.style = 'Table Grid'
    
    table.columns[0].width = Inches(1.5)
    table.columns[1].width = Inches(4.5)
    
    # Helper function to add rows to the detailed vulnerability table
    def add_row_if_data(label, data, impact=None):
        if data and data.strip():  # Only add row if data exists
            row_cells = table.add_row().cells
            row_cells[0].text = label
            set_font(row_cells[0].paragraphs[0].runs[0], bold=True)
            
            data_run = row_cells[1].paragraphs[0].add_run(data)
            if impact:
                data_run.font.color.rgb = get_impact_color(impact)  # Apply color based on impact level
            set_font(data_run)
            set_text_direction_ltr(row_cells[1])

    # Set text direction to LTR
    def set_text_direction_ltr(cell):
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '0')  # Set text direction to LTR
        tcPr.append(bidi)

    impact = first_row['Severity']
    add_row_if_data("Risk", first_row['Severity'] if first_row['Severity'] else 'n/a', impact)
    
    affected_hosts = list(set(affected_hosts))  # Remove duplicates
    affected_hosts_text = ', '.join([f"{host}:{port}" for host, port in affected_hosts])
    add_row_if_data("Affected Hosts", affected_hosts_text if affected_hosts_text else 'n/a')
    
    add_row_if_data("References", first_row['CVE'] if first_row['CVE'] else 'n/a')

    if first_row['Description'] and first_row['Description'].strip():
        description_row = table.add_row().cells
        description_row[0].merge(description_row[1])
        description_row[0].text = "Description"
        set_font(description_row[0].paragraphs[0].runs[0], bold=True)
        description_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        desc_content_row = table.add_row().cells
        desc_content_row[0].merge(desc_content_row[1])
        desc_content_row[0].text = first_row['Description']
        set_font(desc_content_row[0].paragraphs[0].runs[0])
        desc_content_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        desc_content_row[0].paragraphs[0].paragraph_format.space_before = Pt(10)
        desc_content_row[0].paragraphs[0].paragraph_format.space_after = Pt(10)

    if first_row['Steps to Remediate'] and first_row['Steps to Remediate'].strip():
        recommendation_row = table.add_row().cells
        recommendation_row[0].merge(recommendation_row[1])
        recommendation_row[0].text = "Recommendations"
        set_font(recommendation_row[0].paragraphs[0].runs[0], bold=True)
        recommendation_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        recommendation_content_row = table.add_row().cells
        recommendation_content_row[0].merge(recommendation_content_row[1])
        recommendation_content_row[0].text = first_row['Steps to Remediate']
        set_font(recommendation_content_row[0].paragraphs[0].runs[0])
        recommendation_content_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        recommendation_content_row[0].paragraphs[0].paragraph_format.space_before = Pt(10)
        recommendation_content_row[0].paragraphs[0].paragraph_format.space_after = Pt(10)

    # Set table alignment and apply colors to impact cells within the table
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for row in table.rows:
        for cell in row.cells:
            if cell.text in ['Critical', 'High', 'Medium', 'Low']:
                cell_shading = OxmlElement('w:shd')
                cell_shading.set(qn('w:fill'), {
                    'Critical': '8B0000',  # Burgundy
                    'High': 'FF0000',      # Red
                    'Medium': 'FFA500',    # Orange
                    'Low': '006400'        # Dark Green
                }[cell.text])
                cell._element.get_or_add_tcPr().append(cell_shading)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        set_font(run, color=RGBColor(255, 255, 255))  # White text color
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Insert a page break after each vulnerability table
    add_page_break()
    
    vuln_number += 1

# Update Table of Contents
doc.add_paragraph('To update the Table of Contents, select it and press F9.', style='Normal')

# Save the document
doc.save(output_file)

print(f"Summary table, pie chart, and vulnerabilities added to {output_file}.")
