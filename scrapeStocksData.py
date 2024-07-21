import pandas as pd
import matplotlib.pyplot as plt

def load_sheet_names(file_path):
    return pd.ExcelFile(file_path).sheet_names

def read_sheet(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')

def find_TTM_row(df):
    TTM_row = df[df.iloc[:, 0].astype(str).str.contains("TTM", na=False)].index
    if TTM_row.empty:
        TTM_row = df[df.iloc[:, 0].astype(str).str.contains("Breakdown", na=False)].index
    return TTM_row[0] if not TTM_row.empty else None

def extract_x_axis_data(df, TTM_row):
    x_axis_data = []
    for value in df.iloc[TTM_row + 1:, 0]:
        if isinstance(value, str) and not value.isdigit():
            break
        x_axis_data.append(value)
    return pd.Series(x_axis_data).values

def extract_sections(df, TTM_row):
    sections = {}
    current_section = None
    current_data = []

    for i in range(TTM_row + 1, len(df)):
        cell_value = df.iloc[i, 0]
        if pd.isna(cell_value):  # Check for blank cell
            continue
        if isinstance(cell_value, str) and not cell_value.isdigit():
            # If a new section is found, save the current section
            if current_section:
                if len(current_data) > 4:
                    current_data = current_data[-4:]
                elif len(current_data) == 5:
                    current_data = current_data[1:]
                sections[current_section] = current_data
            current_section = cell_value
            current_data = []
        else:
            if cell_value == '--':
                current_data.append(0)
            else:
                current_data.append(cell_value)
    
    if current_section:
        if len(current_data) > 4:
            current_data = current_data[-4:]
        elif len(current_data) == 5:
            current_data = current_data[1:]
        sections[current_section] = current_data

    return sections

def plot_section(sheet_name, section, x_axis, y_axis):
    sheet_prefix = sheet_name.split()[0]  # Extract the first word or letters from the sheet name
    plt.figure()
    plt.plot(x_axis, y_axis, marker='o')
    plt.title(f"{sheet_prefix} {section}")
    plt.xlabel('Date')
    plt.ylabel(f"{sheet_prefix} {section}")
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{sheet_name}_{section.replace(' ', '_')}.png")
    plt.close()

def process_sheet(file_path, sheet_name):
    print(f"Processing sheet: {sheet_name}")
    df = read_sheet(file_path, sheet_name)
    TTM_row = find_TTM_row(df)
    
    if TTM_row is None:
        print(f"Error: 'TTM' or 'Breakdown' not found in sheet '{sheet_name}'. Skipping sheet.")
        print("First two rows in column A:")
        print(df.iloc[0:2, 0].to_string(index=False))
        return
    
    x_axis = extract_x_axis_data(df, TTM_row)
    sections = extract_sections(df, TTM_row)

    for section, y_axis in sections.items():
        if len(x_axis) == len(y_axis):
            plot_section(sheet_name, section, x_axis, y_axis)
        else:
            print(f"Error: Mismatched lengths for section '{section}'. Skipping plot.")
            print(f"Section '{section}' x_axis length: {len(x_axis)}, y_axis length: {len(y_axis)}")

def main(file_path):
    sheet_names = load_sheet_names(file_path)
    for sheet_name in sheet_names:
        process_sheet(file_path, sheet_name)

if __name__ == "__main__":
    file_path = 'Nike.xlsx'
    main(file_path)
