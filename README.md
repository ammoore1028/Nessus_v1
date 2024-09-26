# Vulnerability Assessment Report Automation

This project automates the creation of a vulnerability assessment report from a CSV file of vulnerability data. It converts scan results into a well-formatted Word document that includes:

- A **pie chart** showing the distribution of vulnerabilities based on severity levels (Critical, High, Medium, Low, Info).
- **Tables** for each vulnerability, providing details such as risk level, affected hosts, CVE references, and remediation suggestions.
- Each table is displayed on a **separate page** for cleaner formatting.

## Use Case: Automating the Vulnerability Report

This project is designed to work with results from a vulnerability scan tool, such as the one shown in the image below:

![Image1.png](Image1.png)

The scan output shows a list of vulnerabilities with details like:
- **Severity Level**: Critical, High, Medium, Low, Info
- **CVSS Score**: A score that quantifies the severity of the vulnerability.
- **Vulnerability Name**: The name or type of the vulnerability.
- **Family**: The category or family the vulnerability belongs to.
- **Affected Hosts**: The hosts affected by this vulnerability.
- **Count**: The number of instances detected.

## How It Works

This script reads a CSV file that contains vulnerability data from the scan results. It generates a Word document with:
1. A **pie chart** that visualizes the distribution of vulnerabilities by risk level.
2. A detailed **table for each vulnerability**, including the risk level, affected hosts, CVEs, and remediation suggestions. Each table is placed on a separate page for easier reading and review.

## Features

- **Automated report generation**: Converts CSV data into a structured Word report.
- **Pie chart for visual insights**: Quickly understand the severity distribution.
- **Color-coded risk levels**: Each vulnerability's risk level is color-coded for clarity.
- **Supports multiple hosts**: Lists all affected hosts for each vulnerability.
- **Page breaks for each vulnerability**: Ensures that each vulnerability is presented clearly.

## Installation and Setup

1. **Install Required Python Libraries**

   To use the script, you'll need to install the following libraries:

   ```bash
   pip install python-docx matplotlib
