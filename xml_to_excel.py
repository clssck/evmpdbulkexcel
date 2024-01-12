import os
import sys
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from typing import List, Set
import argparse

# Function to parse XML file and convert it to DataFrame
def parse_xml_to_df(xml_file_path: str, all_columns: Set[str], use_namespace_extraction: bool = False) -> pd.DataFrame:
    try:
        # Parse XML file
        tree = ET.parse(xml_file_path)
    except ET.ParseError:
        # Handle parsing errors
        print(f"Error: Failed to parse {xml_file_path}")
        return pd.DataFrame()

    # Get root of the XML tree
    root = tree.getroot()

    # Nested function to extract data from XML element
    def extract_data(element, parent_path=''):
        data = {}
        for child in element:
            child_path = f"{parent_path}/{child.tag}" if parent_path else child.tag
            if list(child):  # if the child has children
                # Recursively extract data from child
                data.update(extract_data(child, child_path))
            else:
                # Use the tag name of the deepest nested element as the column name
                column_name = child_path.split('/')[-1]
                # Add the text content of the child to the data dictionary
                data[column_name] = child.text.strip() if child.text else None
                # Add the column name to the set of all column names
                all_columns.add(column_name)
        return data

    # List to store all data
    all_data = []
    # Find all 'product' elements in the root and extract data from them
    for product in root.findall('{http://eudravigilance.ema.europa.eu/schema/productExport}product'):
        product_data = extract_data(product)
        all_data.append(product_data)

    # Create DataFrame from all data
    df = pd.DataFrame(all_data)
    # Remove namespace from column names
    df.columns = [col.split('}')[-1] for col in df.columns]
    return df
    # Function to process multiple XML files
def process_xml_files(xml_files: List[str], input_folder: str, use_namespace_extraction: bool = False) -> pd.DataFrame:
    all_data = []
    total_products = 0
    all_columns = set()  # create a set to store all column names

    for xml_file in xml_files:
        xml_file_path = os.path.join(input_folder, xml_file)
        df = parse_xml_to_df(xml_file_path, all_columns, use_namespace_extraction)
        total_products += len(df)
        all_data.append(df)

    # Concatenate all DataFrames
    final_df = pd.concat(all_data)
    return final_df, total_products, len(all_columns)

# Function to save DataFrame to Excel file
def save_to_excel(df: pd.DataFrame, output_excel_file_path: str) -> None:
    try:
        # Save DataFrame to Excel file
        df.to_excel(output_excel_file_path, index=False)
        # Load the workbook
        wb = load_workbook(output_excel_file_path)
        # Get the active sheet
        sheet = wb.active
        # Set the width of all columns to 30
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = 30
        # Create a table style
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        # Create a table
        table = Table(displayName="Table1", ref=sheet.dimensions)
        # Apply the table style
        table.tableStyleInfo = style
        # Add the table to the sheet
        sheet.add_table(table)
        # Save the workbook
        wb.save(output_excel_file_path)
    except Exception as e:
        # Handle exceptions when saving Excel file
        print(f"Error: Failed to save Excel file {output_excel_file_path}. {e}")

# Function to process XML files in a folder
def process_xml_folder(input_folder: str, output_folder: str = None) -> None:
    # Check if the input folder exists
    if not os.path.exists(input_folder):
        print(f"Error: The folder {input_folder} does not exist.")
        return

    # Get a list of all XML files in the input folder
    xml_files = [f for f in os.listdir(input_folder) if f.endswith('.xml')]

    # Ask the user if they want to use namespace extraction
    use_namespace_extraction = input("Do you want to use namespace extraction? (yes/no): ").lower() == 'yes'

    # Process each XML file and add the data to a DataFrame
    final_df, total_products, num_columns = process_xml_files(xml_files, input_folder, use_namespace_extraction)

    # Save the DataFrame to an Excel file
    output_folder = output_folder or input_folder
    output_excel_file_path = os.path.join(output_folder, f"output_{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx")
    save_to_excel(final_df, output_excel_file_path)

    # Check if the file was saved successfully
    if os.path.exists(output_excel_file_path):
        print(f"Successfully processed {len(xml_files)} XML files and saved the data to {output_excel_file_path}.")
    else:
        print(f"Error: Failed to save Excel file {output_excel_file_path}.")

    # Print the number of XML files parsed
    print(f"Parsed {len(xml_files)} XML files.")

    # Print the number of products and unique columns
    print(f"You have {total_products} products with {num_columns} unique columns.")

    # Ask the user if they want to proceed
    proceed = input("Do you want to proceed? (yes/no): ")
    if proceed.lower() != 'yes':
        print("Operation cancelled.")
        return

    # Generate the output Excel file path
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_excel_file_path = os.path.join(input_folder, f"output_{timestamp}.xlsx")

    # Print a success message
    print(f"Successfully processed {len(xml_files)} XML files and saved the data to {output_excel_file_path}.")

if __name__ == "__main__":
    # Create the parser
    parser = argparse.ArgumentParser(description="Process XML files in a folder and save the data to an Excel file.")

    # Add the arguments
    parser.add_argument('input_folder', type=str, help='The path to the input folder containing the XML files. You can use "." to use the directory where the script is located.')
    parser.add_argument('-o', '--output_folder', type=str, help='The path to the output folder where the Excel file will be saved. If not provided, the Excel file will be saved in the input folder.')

    # Parse the arguments
    args = parser.parse_args()

    # If the input folder is ".", use the directory of the script
    if args.input_folder == ".":
        input_folder = os.path.dirname(os.path.realpath(__file__))
    else:
        # Convert the input folder path to an absolute path
        input_folder = os.path.abspath(args.input_folder)

    # If the output folder is provided, convert it to an absolute path
    output_folder = os.path.abspath(args.output_folder) if args.output_folder else None

    # Run the main function with the parsed arguments
    process_xml_folder(input_folder, output_folder)