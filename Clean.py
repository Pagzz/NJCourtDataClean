import pandas as pd
import re

def load_dnc_list(dnc_file_path):
    """Load the Do Not Call list from an Excel file with multiple tags."""
    dnc_df = pd.read_excel(dnc_file_path)
    # Convert all relevant columns to lowercase for case-insensitive comparison
    for column in ['Last_Name', 'First_Name']:
        dnc_df[column] = dnc_df[column].str.lower()
    return dnc_df



# Define the set of violation codes to filter on
included_violation_codes = {
    '39:4-98', '39:4-96', '39:4-50', '39:4-128.1', '39:4-89', '39:4-85', '39:4-86', '39:4-50.15B',
    '39:4-51B', '39:4-50.2', '39:4-51A', '2C:12-1A(1)', '2C:12-1A(3)', '2C:12-1B(1)', '2C:12-1B(2)',
    '2C:12-1B(5)(A)', '2C:12-1B(5)(H)', '2C:12-1B(7)', '2C:12-1C(2)', '2C:12-3.1A(2)', '2C:12-3A',
    '2C:12-3B', '2C:14-2C(4)', '2C:14-4A', '2C:15-1A(1)', '2C:17-3A(1)', '2C:18-3A', '2C:18-3B',
    '2C:18-3B(1)', '2C:20-10A', '2C:20-11B(1)', '2C:20-11B(2)', '2C:20-3A', '2C:20-4', '2C:20-6', '2C:20-7A',
    '2C:20-8A', '2C:20-8C(2)', '2C:21-5', '2C:29-1A', '2C:29-3B(4)', '2C:33-2A(2)', '2C:33-2B', '2C:33-4A',
    '2C:33-4C', '2C:35-10C', '2C:35-5B(11)(A)', '2C:36-2A', '2C:29-3B(4)', '2C:20-11B(5)', '2C:28-7A(1)', '2C:28-7A(1)',
    '2C:29-3B(4)', '2C:29-1A', '2C:29-2A(1)', '2C:29-3A(7)', '2C:29-3B(1)', '2C:33-2A(1)', '2C:33-2A(2)', 
    '2C:33-4A', '2C:33-4B', '2C:33-4C', '2C:34-1B(1)', '2C:35-10A(1)', '2C:35-10C',	
}

def parse_court_info(line):
    """Extract court information from a given line."""
    pattern = r'MUNICIPAL COURT : (\d+)\s+(.*)'
    match = re.search(pattern, line)
    if match:
        court_code, court_name = match.groups()
        return {'Court_Code': court_code, 'Court_Name': court_name.strip()}
    return None


def is_on_dnc_list(record, dnc_df):
    """Check if a record matches any entry on the DNC list based on multiple tags."""
    # Example logic: Check if any record matches by last and first name, email, or phone
    matches = dnc_df[
        (dnc_df['Last_Name'] == record['Last_Name']) & 
        (dnc_df['First_Name'] == record['First_Name'])
    ]
    return not matches.empty

def parse_record(lines, court_info, dnc_df):
    """Parse lines of a record and append court information."""
    violations = [line[117:131].strip() for line in lines if line[117:131].strip() in included_violation_codes]

    if not violations:  # If no included violation codes are found, skip this record
        return None

    record = {
        'Last_Name': lines[0][59:69].strip(),
        'First_Name': lines[0][42:57].strip(),
        'Middle_Initial': lines[0][57:58].strip(),
        'Offense Date': lines[0][18:28].strip(),
        'Issue Date': lines[0][30:40].strip(),
        'Court Date': lines[0][83:93].strip(),
        'Physical_Address': lines[1][42:75].strip(),
        'Physical_City': lines[3][42:63].strip(),
        'Physical_State': lines[3][63:66].strip(),
        'Physical_Zip': lines[3][68:79].strip(),
        **court_info,
        'Violations': ', '.join(violations),
    }

    if is_on_dnc_list(record, dnc_df):
        return None  # Skip this record if it matches the DNC list

    return record

def parse_file_to_df(file_path, dnc_df):
    records = []
    court_info = {}
    with open(file_path, 'r') as file:
        lines_for_record = []
        for line in file:
            if 'MUNICIPAL COURT :' in line:
                court_info = parse_court_info(line)
            elif any(line.startswith(code) for code in ['0S', '0SC']):  # Record start indicators
                if lines_for_record:  # If there are collected lines, process them as a record
                    record = parse_record(lines_for_record, court_info, dnc_df)  # Pass dnc_df here
                    if record:
                        records.append(record)
                    lines_for_record = []  # Reset for a new record
                lines_for_record.append(line)
            elif lines_for_record:  # If already collecting a record, continue adding lines
                lines_for_record.append(line)
        
        # Process the last collected record
        if lines_for_record:
            record = parse_record(lines_for_record, court_info, dnc_df)  # Pass dnc_df here
            if record:
                records.append(record)

    return pd.DataFrame(records)



# Load the DNC list
dnc_file_path = 'C:\\Users\\rp122\\OneDrive\\Desktop\\dataUSA\\Excel\\DNC.xlsx'
dnc_df = load_dnc_list(dnc_file_path)

# Specify the path to your data file
file_path = 'C:\\Users\\rp122\\OneDrive\\Desktop\\dataUSA\\textFiles\\pawcmc0081.txt'
# Now, pass the dnc_df when calling parse_file_to_df
df = parse_file_to_df(file_path, dnc_df)

# Save the filtered data to Excel
output_path = 'C:\\Users\\rp122\\OneDrive\\Desktop\\dataUSA\\Excel\\cleaned_data.xlsx'
df.to_excel(output_path, index=False, engine='openpyxl')
print(f"Data successfully written to {output_path}")