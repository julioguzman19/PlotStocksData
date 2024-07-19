import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = 'Nike.xlsx'
sheet_name = 'NKE Balance Annual'

# Read the Excel sheet into a DataFrame, specifying the openpyxl engine
df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')

# Find the row index where "Breakdown" is located
breakdown_row = df[df.iloc[:, 0].str.contains("Breakdown", na=False)].index[0]

# Extract the x-axis data dynamically, starting from the row after "Breakdown"
x_axis_data = []
for value in df.iloc[breakdown_row + 1:, 0]:
    if isinstance(value, str) and not value.isdigit():
        break
    x_axis_data.append(value)

# Convert x_axis_data to a numpy array
x_axis = pd.Series(x_axis_data).values

# Initialize an empty dictionary to hold sections
sections = {}

# Iterate through the DataFrame starting from the row after "Breakdown"
current_section = None
current_data = []

for i in range(breakdown_row + 1, len(df)):
    cell_value = df.iloc[i, 0]
    if pd.isna(cell_value):  # Check for blank cell
        continue
    if isinstance(cell_value, str) and not cell_value.isdigit():
        # If a new section is found, save the current section
        if current_section:
            sections[current_section] = current_data
        current_section = cell_value
        current_data = []
    else:
        # Append data to the current section
        current_data.append(cell_value)

# Add the last section if any
if current_section:
    sections[current_section] = current_data

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
