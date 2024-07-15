import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = 'Nike.xlsx'  # Ensure the file is in the same directory as your script
sheet_name = 'NKE Balance Annual'  # The sheet name to read from

# Read the Excel file
data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

# Parse the data into a structured DataFrame
data.columns = ['Values']
sections = ['Total Assets']
parsed_data = {section: [] for section in sections}

# # Extract the breakdown dates and values for Total Assets
# current_section = None
# dates = []
# for index, row in data.iterrows():
#     value = row['Values']
#     if value in sections:
#         current_section = value
#     elif isinstance(value, str) and value.startswith('5/31'):
#         dates.append(value)
#     elif current_section and isinstance(value, str) and value.replace(',', '').isdigit():
#         parsed_data[current_section].append(int(value.replace(',', '')))

# # Create a DataFrame for plotting
# df = pd.DataFrame({
#     'Date': dates,
#     'Total Assets': parsed_data['Total Assets']
# })

# # Plot the data
# plt.figure(figsize=(10, 5))
# plt.plot(df['Date'], df['Total Assets'], marker='o')
# plt.title('Total Assets')
# plt.xlabel('Date')
# plt.ylabel('Value')
# plt.tight_layout()
# plt.show()
