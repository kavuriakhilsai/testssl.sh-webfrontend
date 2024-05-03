import os
import json
from docx import Document
from colorama import Fore, Style
import string

# Get the directory of the Flask application file
app_dir = os.path.dirname(__file__)

# Define the path to the JSON result folder
result_folder_path = os.path.join(app_dir, "result", "json")

# Path to output Word file
output_file_path = os.path.join(app_dir, "output.docx")

# Function to filter severity levels
def filter_severity(data):
    severity = data.get("severity", "").upper()
    return severity in ["MEDIUM", "WARN"]

# Function to create color-coded text
def color_text(severity, text):
    if severity == "HIGH":
        return f"{Fore.RED}{text}{Style.RESET_ALL}"
    elif severity == "CRITICAL":
        return f"{Fore.MAGENTA}{text}{Style.RESET_ALL}"
    elif severity == "MEDIUM":
        return f"{Fore.YELLOW}{text}{Style.RESET_ALL}"
    elif severity == "WARN":
        return f"{Fore.BLUE}{text}{Style.RESET_ALL}"
    else:
        return text

# Function to sanitize text for XML compatibility
def sanitize_text(text):
    # Filter out non-printable ASCII and non-ASCII characters
    sanitized_text = "".join(char for char in text if char.isprintable() and ord(char) < 128)
    return sanitized_text

# Create a new Word document
doc = Document()

# Loop through each file in the JSON result folder
for filename in os.listdir(result_folder_path):
    if filename.endswith(".json"):
        file_path = os.path.join(result_folder_path, filename)
        with open(file_path, "r") as file:
            try:
                data_list = json.load(file)
                for data in data_list:
                    if filter_severity(data):
                        # Add severity and finding to the Word document
                        severity = data.get("severity", "Unknown")
                        finding = data.get("finding", "No finding")
                        sanitized_finding = sanitize_text(finding)
                        doc.add_paragraph(f"{severity}: {color_text(severity, sanitized_finding)}")
            except json.JSONDecodeError:
                print(f"Error reading JSON from file: {filename}")

# Save the Word document
doc.save(output_file_path)
