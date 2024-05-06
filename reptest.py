import pandas as pd
import matplotlib.pyplot as plt

#Reading the generated CSV file
df = pd.read_csv("/home/ubuntu/TestSSL/testssl.sh-webfrontend/output_csv/output.csv")

# Calculating the ratio of severieties
severity_counts = df['severity'].value_counts()
total_severities = severity_counts.sum()
severity_ratios = severity_counts / total_severities

# Create a bar chart
plt.figure(figsize=(10, 6))
severity_ratios.plot(kind='bar', color='skyblue')
plt.title('Severity Ratios')
plt.xlabel('Severity')
plt.ylabel('Ratio')
plt.xticks(rotation=45)
plt.tight_layout()

# Show the plot
plt.show()
