import json
import os
import csv

# JSON file Path
json_file_path = ""

# Path to CSV folder
csv_folder_path = ""
#Creating CSV folder if it doesnt exist
os.makedirs(csv_ folder_path, exist_ok = True)

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
