import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = 'Nike.xlsx'
sheet_name = 'NKE Balance Annual'

# Read the Excel sheet into a DataFrame, specifying the openpyxl engine
df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')

# Extract the x-axis data from cells A2-A5
x_axis = df.iloc[1:5, 0].values

# Define the sections and their corresponding y-axis data
sections = {
    "Total Assets": df.iloc[6:10, 0].values,
    "Total Liabilities Net Minority Interest": df.iloc[11:15, 0].values,
    "Total Equity Gross Minority Interest": df.iloc[16:20, 0].values,
    "Total Capitalization": df.iloc[21:25, 0].values,
    "Common Stock Equity": df.iloc[26:30, 0].values
}

# Plotting each section with "Nike" included in the titles and y-axis labels
for section, y_axis in sections.items():
    plt.figure()
    plt.plot(x_axis, y_axis, marker='o')
    plt.title(f"Nike {section}")
    plt.xlabel('Date')
    plt.ylabel(f"Nike {section}")
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{section.replace(' ', '_')}.png")
    plt.show()
