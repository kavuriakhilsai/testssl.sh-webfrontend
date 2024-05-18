import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
import os

# Define paths
app_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(app_dir, "output_csv", "output.csv")
plot_image_path = os.path.join(app_dir, "severity_plot.png")
word_doc_path = os.path.join(app_dir, "severity_report.docx")

# Read the CSV into a DataFrame
df = pd.read_csv("/home/ubuntu/TestSSL/testssl.sh-webfrontend/output_csv/output.csv")

# Plot the severities
severity_counts = df['severity'].value_counts()
plt.figure(figsize=(10, 6))
severity_counts.plot(kind='bar', color=['blue', 'orange', 'red', 'green'])
plt.title('Severity Counts')
plt.xlabel('Severity')
plt.ylabel('Count')
plt.xticks(rotation=0)
plt.tight_layout()

# Save the plot
plt.savefig(plot_image_path)
plt.close()

# Create a Word document
doc = Document()
doc.add_heading('Severity Report', 0)

# Add plot image to the Word document
doc.add_picture(plot_image_path, width=Inches(6))

# Save the Word document
doc.save(word_doc_path)

print(f"Word document created and saved at {word_doc_path}")

