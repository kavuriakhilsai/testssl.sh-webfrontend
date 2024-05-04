import json
import os
import csv

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

# Path to CSV folder
csv_folder_path = os.path.join(app_dir, "output.csv")

#Creating CSV folder if it doesnt exist
os.makedirs(csv_folder_path, exist_ok = True)

#Path ti the output CSV file
csv_file_path = os.path.join(csv_folder_path, "output.csv")

# Opening the JSON file and loading the data
with open(json_file_path, "r") as json_file:
  data = json.load(json_file)

# Define the field names for the CSV file
field_names = ["id", "ip", "port", "severity","cve","cwe","finding"]

# Open the CSV file and write the data
with open(csv_file_path, "w", newline="") as csv_file:
  writer = csv.DictWriter(csv_file, field_names = field_names)
  writer.writeheader()
  for item in data:
    # Checking if the "cve" or "CWE" keys are present and provide default values if not available
    cve = item.get("cve","N/A")
    cwe = item.get("cwe","N/A")
    row = {"id":item["id"],"ip":item["ip"],"port":item["port"],"severity":item["severity"],"finding":item["finding"],"cve":cve, "cwe":cwe}
  writer.writerow(row)
