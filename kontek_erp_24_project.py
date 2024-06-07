import pandas as pd
import json

def parse_bom_excel(input_file):
    xl = pd.ExcelFile(input_file)
    df = xl.parse(header=None, skiprows=1)  

    parts = {}
    assemblies = {}

    for index, row in df.iterrows():
        item_no = str(row[0]).strip() if pd.notna(row[0]) else ''
        part_number = str(row[1]).strip() if pd.notna(row[1]) else ''
        description = str(row[2]).strip() if pd.notna(row[2]) else ''
        sw_file_name = str(row[3]).strip() if pd.notna(row[3]) else ''
        quantity = row[5] if pd.notna(row[5]) else 0

        part_number = part_number.strip()

        if part_number in parts:
            parts[part_number]['item_no'].append(item_no)
            parts[part_number]['quantities'].append(quantity)
            parts[part_number]['sum_total_qty'] += quantity
        else:
            parts[part_number] = {
                'item_no': [item_no],
                'description': description,
                'sw_file_name': sw_file_name,
                'sum_total_qty': quantity,
                'quantities': [quantity]
            }

        print(f"Logged part: {part_number}")

    return parts, assemblies

def save_to_json(data, filename):
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

def main():
    input_file = "P:/KONTEK/ENGINEERING/ELECTRICAL/Application Development/Spreadsheets to Parse/2023699.xls"
    parts, assemblies = parse_bom_excel(input_file)

    save_to_json(parts, 'parts.json')
    save_to_json(assemblies, 'assemblies.json')

    print("\nParsing Complete!\n")
    print(f"Logged {len(parts)} parts to parts.json")
    print(f"Logged {len(assemblies)} assemblies to assemblies.json")

if __name__ == "__main__":
    main()

