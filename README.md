# Vulnerability Report Generator Script

This Python script generates a vulnerability assessment report based on a CSV file input. It creates a Word document that includes a pie chart of the vulnerabilities and tables detailing each vulnerability. Each table starts on a new page, and the pie chart is placed on the first page of the document.

## Features

- **Pie Chart Generation**: The script creates a pie chart to visualize the distribution of vulnerabilities by risk level (Critical, High, Medium, Low).
- **Word Document Creation**: It generates a Word document using `python-docx` that includes a detailed table for each vulnerability.
- **Color-coded Risk Levels**: The risk levels are color-coded (Critical = Burgundy, High = Red, Medium = Orange, Low = Dark Green).
- **Page Breaks**: Each vulnerability is displayed in a separate table on a new page in the document, ensuring clean and organized formatting.
- **Supports CSV Input**: The script processes a CSV file with vulnerability details, making it adaptable to various vulnerability scanning tools' output.

## Libraries Required

To run this script, you'll need to install the following Python libraries:

- `csv`: To parse the CSV file containing the vulnerability data (built-in).
- `python-docx`: To create and format the Word document.
- `matplotlib`: To generate and save the pie chart for vulnerabilities.
- `argparse`: To handle command-line arguments for CSV file input (built-in).
- `collections`: To use `defaultdict` for grouping vulnerabilities (built-in).
- `os`: To handle file operations (built-in).

### Installing the Required Libraries

To install the required libraries, you can use `pip` by running the following command:

```bash
pip install python-docx matplotlib

