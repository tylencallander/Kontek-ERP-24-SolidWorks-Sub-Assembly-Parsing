import pandas as pd
import json
import os

def parse_bom_excel(input_file):
    xl = pd.ExcelFile(input_file)
    df = xl.parse(header=None, skiprows=1) 

    parts = {}
    errors = {"EMPTYCELL": []}

    for index, row in df.iterrows():
        item_no = str(row[0]).strip() if pd.notna(row[0]) else ''
        part_number = str(row[1]).strip() if pd.notna(row[1]) else ''
        description = str(row[2]).strip() if pd.notna(row[2]) else ''
        sw_file_name = str(row[3]).strip() if pd.notna(row[3]) else ''
        quantity = row[5] if pd.notna(row[5]) else 0

# I didn't really have any error conditions to make, since im not comparing this excel sheet to the network like I was doing before, so I just logged items with empty cells to the error file.
# I was going to also log items that have a newline in their title, since it kinda ruins the formatting of the part number, ex: "STW 0.3125-\n18x1.75 SS" but I chose not to, but I can add this if necessary.

        if not part_number or not item_no or not description or not sw_file_name or not quantity:
            errors['EMPTYCELL'].append({
                'row': index + 2, 
                'item_no': item_no,
                'part_number': part_number,
                'description': description,
                'sw_file_name': sw_file_name,
                'quantity': quantity
            })
            print(f"Error logged for empty cells at row {index + 2}")
            continue

        if part_number in parts:
            parts[part_number]['item_no'].append(item_no)
            parts[part_number]['quantities'].append(quantity)
            parts[part_number]['total_qty'] += quantity
        else:
            parts[part_number] = {
                'item_no': [item_no],
                'description': description,
                'SW_file_name': sw_file_name,
                'total_qty': quantity,
                'quantities': [quantity]
            }

        print(f"Processed part: {part_number}")

    return parts, errors, df

def build_assemblies(df):
    assemblies = {}
    stack = []

    for _, row in df.iterrows():
        item_no = str(row[0]).strip() if pd.notna(row[0]) else ''
        part_number = str(row[1]).strip() if pd.notna(row[1]) else ''
        description = str(row[2]).strip() if pd.notna(row[2]) else ''
        sw_file_name = str(row[3]).strip() if pd.notna(row[3]) else ''

        if not item_no or not part_number or not description or not sw_file_name:
            continue

        part_number = part_number.strip()

        level = item_no.count('.')
        assembly = {
            "item_no": item_no,
            "description": description,
            "SW_file_name": sw_file_name,
            "assemblies": []
        }

# Used a stack to figure out the parent assemblies

        while len(stack) > level:
            stack.pop()

        if stack:
            parent = stack[-1]
            parent["assemblies"].append({part_number: assembly})
        else:
            assemblies[part_number] = assembly

        stack.append(assembly)

    return assemblies

def save_to_json(data, filename):
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

def main():
    input_file = "P:/KONTEK/ENGINEERING/ELECTRICAL/Application Development/Spreadsheets to Parse/2023699.xls"
    project_name = os.path.splitext(os.path.basename(input_file))[0]

    parts, errors, df = parse_bom_excel(input_file)
    assemblies = build_assemblies(df)

    save_to_json(parts, 'parts.json')
    save_to_json(errors, 'errors.json')

    assembly_structure = {
        "project": project_name,
        "name": "",
        "assemblies": assemblies
    }
    save_to_json(assembly_structure, 'assemblies.json')

    print("\nParsing Complete!\n")
    print(f"Logged {len(parts)} parts to parts.json")
    print(f"Logged {len(errors['EMPTYCELL'])} errors to errors.json")

# This logs the main parent assemblies, like 1.0, 2.0... Not all of the individual parts that make up the assemblies.
    print(f"Logged {len(assemblies)} parent assemblies to assemblies.json")

if __name__ == "__main__":
    main()