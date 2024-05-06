import pandas as pd
import matplotlib.pyplot as plt

#Reading the generated CSV file
df = pd.read_csv("/home/ubuntu/TestSSL/testssl.sh-webfrontend/output_csv/output.csv")

# Calculating the ratio of severieties
severity_counts = df['severity'].value_counts()
total_severities = severity_counts.sum()
severity_ratios = severity_counts / total_severities

print(severity_counts)
warn_severity = df[df['severity'] == 'WARN']
print(warn_severity)
