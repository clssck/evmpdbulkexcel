# evmpdbulkexcel
Used for extracting all data from EVMPD BULK EXPORT xml's to generate an excel table


This Python script is designed to convert XML files to an Excel file. It provides a command-line interface to specify the input folder containing the XML files and an optional output folder to save the generated Excel file. If no output folder is provided, the Excel file is saved in the input folder.

The script contains several functions:

parse_xml_to_df: Parses an XML file and converts it to a pandas DataFrame. It also handles XML parsing errors.

process_xml_files: Processes multiple XML files in a given folder and concatenates the resulting DataFrames.

save_to_excel: Saves a DataFrame to an Excel file and formats it as a table with a specified style. It also handles exceptions when saving the Excel file.

process_xml_folder: Processes all XML files in a given folder and saves the data to an Excel file. It checks if the input folder exists, gets a list of all XML files in the folder, processes each XML file, and saves the DataFrame to an Excel file.

The script uses the argparse module to parse command-line arguments for the input and output folders. It runs the process_xml_folder function with the parsed arguments when the script is run.
