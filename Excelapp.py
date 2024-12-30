import configparser
import pandas as pd
import re

def parse_condition(column_name, condition, df):
    if "between" in condition:
        match = re.search(r'between\s+(\d+)\s+and\s+(\d+)', condition)
        if match:
            start, end = map(float, match.groups())
            return df[(df[column_name] >= start) & (df[column_name] <= end)]
    elif "in" in condition:
        match = re.search(r'in\s+\[(.*?)\]', condition)
        if match:
            values = match.group(1).split(',')
            values = [v.strip().strip('"').strip("'") for v in values]
            # Convert values to appropriate types if needed (numeric for 'Salary' or other numeric columns)
            if df[column_name].dtype in ['int64', 'float64']:
                values = [float(v) for v in values]
            return df[df[column_name].isin(values)]
    elif "=" in condition:
        match = re.search(r'=\s*(.+)', condition)
        if match:
            value = match.group(1).strip().strip('"').strip("'")
            # Convert to numeric if needed
            if df[column_name].dtype in ['int64', 'float64']:
                value = float(value)
            return df[df[column_name] == value]
    else:
        print(f"Unsupported condition: {condition}. Skipping.")
    return df

def process_excel(input_excel, ini_file, output_excel):
    # Parse the INI file
    config = configparser.ConfigParser()
    config.read(ini_file)
    
    # Read sections from the INI file
    required_columns = [col.strip().strip('"') for col in config['Columns']['required_columns'].split(',')]
    filters = dict(config['Filters'])
    
    # Read the Excel file
    df = pd.read_excel(input_excel)
    
    # Select required columns
    df = df[required_columns]
    # Apply filtering conditions
    for column_with_quotes, condition in filters.items():
        column_name = column_with_quotes.strip('"')
        # Find the actual column name in the DataFrame (case-insensitive match)
        matching_columns = [col for col in df.columns if col.lower() == column_name.lower()]
        
        if matching_columns:
            column_name = matching_columns[0]
            df = parse_condition(column_name, condition, df)
        else:
            print(f"Column {column_name} not found in the XLSX file. Skipping.")
    
    # Write the filtered data to a new Excel file
    df.to_excel(output_excel, index=False)
    print(f"Filtered data saved to {output_excel}")

# Example usage
process_excel('Input.xlsx', 'Rules.ini', 'Output.xlsx')