import os
import sys
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
from openpyxl import load_workbook

def parse_xml_to_df(xml_file_path, all_columns):
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    def extract_data(element, parent_path=''):
        data = {}
        for child in element:
            child_path = f"{parent_path}/{child.tag}" if parent_path else child.tag
            if list(child):  # if the child has children
                data.update(extract_data(child, child_path))
            else:
                # Use the tag name of the deepest nested element as the column name
                column_name = child_path.split('/')[-1]
                data[column_name] = child.text.strip() if child.text else None
                all_columns.add(column_name)  # add the column name to the set
        return data

    all_data = []
    for product in root.findall('{http://eudravigilance.ema.europa.eu/schema/productExport}product'):
        product_data = extract_data(product)
        all_data.append(product_data)

    df = pd.DataFrame(all_data)
    df.columns = [col.split('}')[-1] for col in df.columns]  # remove namespace from column names
    return df

def process_xml_folder(input_folder):
    xml_files = [f for f in os.listdir(input_folder) if f.endswith('.xml')]
    all_data = []
    total_products = 0
    all_columns = set()  # create a set to store all column names

    for xml_file in xml_files:
        xml_file_path = os.path.join(input_folder, xml_file)
        df = parse_xml_to_df(xml_file_path, all_columns)
        total_products += len(df)
        all_data.append(df)

    final_df = pd.concat(all_data)
    num_columns = len(all_columns)  # get the number of unique columns

    print(f"You have {total_products} products with {num_columns} unique columns about to be added to the table.")
    proceed = input("Do you want to proceed? (yes/no): ")
    if proceed.lower() != 'yes':
        print("Operation cancelled.")
        return

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_excel_file_path = os.path.join(input_folder, f"output_{timestamp}.xlsx")

    # Use openpyxl to create Excel file and set column width
    final_df.to_excel(output_excel_file_path, index=False)
    wb = load_workbook(output_excel_file_path)
    sheet = wb.active
    for column in sheet.columns:
        sheet.column_dimensions[column[0].column_letter].width = 30
    wb.save(output_excel_file_path)

    print(f"Data extracted to {output_excel_file_path}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python xml_to_excel.py <input_folder>")
        sys.exit(1)

    input_folder = sys.argv[1]
    process_xml_folder(input_folder)