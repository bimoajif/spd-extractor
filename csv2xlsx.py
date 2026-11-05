import pandas as pd

# Load the CSV file
df = pd.read_csv("./results/hasil_ekstraksi.csv")

# Save it as an Excel file
df.to_excel("output_file.xlsx", index=False)