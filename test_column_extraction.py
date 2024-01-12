import os
import sys
import pandas as pd
import xml.etree.ElementTree as ET
from lxml import etree
from collections import defaultdict
import time

def parse_xml_to_df_etree(xml_file_path):
    tree = ET.parse(xml_file_path)
    root = tree.getroot()
    return root

def parse_xml_to_df_lxml(xml_file_path):
    tree = etree.parse(xml_file_path)
    root = tree.getroot()
    return root

def extract_data_from_root_1(root):
    def extract_data(element, parent_path=''):
        data = defaultdict(str)
        for child in element:
            child_path = f"{parent_path}/{child.tag}" if parent_path else child.tag
            if list(child):  # if the child has children
                data.update(extract_data(child, child_path))
            else:
                column_name = child_path.split('/')[-1]
                data[column_name] = child.text.strip() if child.text else None
        return data

    all_data = []
    for product in root.findall('{http://eudravigilance.ema.europa.eu/schema/productExport}product'):
        product_data = extract_data(product)
        all_data.append(product_data)

    df = pd.DataFrame(all_data)
    df.columns = [col.split('}')[-1] for col in df.columns]  # remove namespace from column names
    return df

def extract_data_from_root_2(root):
    def extract_data(element, parent_path=''):
        data = defaultdict(str)
        for child in element:
            child_path = f"{parent_path}/{child.tag}" if parent_path else child.tag
            if list(child):  # if the child has children
                data.update(extract_data(child, child_path))
            else:
                data[child_path] = child.text.strip() if child.text else None
        return data

    all_data = []
    for product in root.findall('{http://eudravigilance.ema.europa.eu/schema/productExport}product'):
        product_data = extract_data(product)
        all_data.append(product_data)

    df = pd.DataFrame(all_data)
    df.columns = [col.split('}')[-1] for col in df.columns]  # remove namespace from column names
    return df

def extract_data_from_root_3(root):
    all_data = []
    for product in root.findall('{http://eudravigilance.ema.europa.eu/schema/productExport}product'):
        product_data = {child.tag: child.text.strip() if child.text else None for child in product}
        all_data.append(product_data)

    df = pd.DataFrame(all_data)
    df.columns = [col.split('}')[-1] for col in df.columns]  # remove namespace from column names
    return df

def process_xml_folder(input_folder, parse_xml_to_df, extract_data_from_root):
    xml_files = [f for f in os.listdir(input_folder) if f.endswith('.xml')]

    column_counts = defaultdict(int)

    start_time = time.time()
    for xml_file in xml_files:
        xml_file_path = os.path.join(input_folder, xml_file)
        root = parse_xml_to_df(xml_file_path)
        df = extract_data_from_root(root)

        for column in df.columns:
            column_counts[column] += 1
    end_time = time.time()

    processing_time = end_time - start_time
    unique_columns = len(column_counts)

    return processing_time, unique_columns

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python test_column_extraction.py <input_folder>")
        sys.exit(1)

    input_folder = sys.argv[1]

    methods = [
        ("xml.etree.ElementTree", parse_xml_to_df_etree),
        ("lxml.etree", parse_xml_to_df_lxml)
    ]

    extractions = [
        ("extraction logic 1", extract_data_from_root_1),
        ("extraction logic 2", extract_data_from_root_2),
        ("extraction logic 3", extract_data_from_root_3)
    ]

    for method_name, method in methods:
        for extraction_name, extraction in extractions:
            print(f"Processing with {method_name} and {extraction_name}:")
            processing_time, unique_columns = process_xml_folder(input_folder, method, extraction)
            print(f"Processing took {processing_time} seconds.")
            print(f"Found {unique_columns} unique columns across all files.\n")