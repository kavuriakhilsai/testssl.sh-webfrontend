import osimport json
from docx import Document
from colorama import Fore, Style

#Getting the directory of the Flask Application File
app_dir = os.path.dirname(__file__)

#Defining the path to the JSON folder
result_folder_path = os.path.join(app_dir, "result", "json")

#Output word file path
output_file_path = os.path.join(app_dir, "output.docx")

# Function to filter severity levels
def filter_severity(data):
    severity = data.get("severity", "").upper()
    return severity in ["HIGH", "CRITICAL", "MEDIUM", "WARN"]

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
    # Create a new Word document
doc = Document()

# Loop through each file in the JSON result folder
for filename in os.listdir(result_folder_path):
    if filename.endswith(".json"):
        file_path = os.path.join(result_folder_path, filename)
        with open(file_path, "r") as file:
            try:
                data = json.load(file)
                if filter_severity(data):
                    # Add severity and finding to the Word document
                    severity = data.get("severity", "Unknown")
                    finding = data.get("finding", "No finding")
                    doc.add_paragraph(f"{severity}: {color_text(severity, finding)}")
            except json.JSONDecodeError:
                print(f"Error reading JSON from file: {filename}")

# Save the Word document
doc.save(output_file_path)