import os
import csv
import pandas as pd
from datetime import datetime, timedelta
import re
import shutil
import logging

def from_excel_ordinal(ordinal: float, _epoch0=datetime(1899, 12, 31)) -> datetime:
    if ordinal >= 60:
        ordinal -= 1  # Excel leap year bug, 1900 is not a leap year!
    return (_epoch0 + timedelta(days=ordinal)).replace(microsecond=0)

# Ruta Salida Archivo
OUTPUT_FILE_PATH = "C:/Users/Public/Documents/Reporte Oxidos - Oficial/Input"

# Ruta Archivo Log
LOG_FILE_PATH = "C:/Phyton Scripts/Excel to CSV/log/Log.txt"

# Function to read input parameters
def read_input_parameters(input_csv_path):
    try:
        with open(input_csv_path, mode="r") as file:
            csv_reader = csv.DictReader(file)
            parameters = [
                {key.strip(): value.strip() for key, value in row.items()}
                for row in csv_reader
            ]
            return parameters
    except Exception as e:
        print(f"Error reading input parameters: {e}")
        logger.error(e,'%s Error reading input parameters %s')
        return []

# Function to extract timestamp from the Excel filename
def extract_timestamp_from_filename(filename):
    try:
        # Remove the .xlsx extension before matching the pattern
        base_filename = os.path.splitext(filename)[0]
        # Match the pattern "dd MM yyyy" at the end of the filename
        match = re.search(r'(\d{2}) (\d{2}) (\d{4})$', base_filename)
        if match:
            day, month, year = map(int, match.groups())
            return datetime(year, month, day, 7, 0, 0)  # 7 AM of the extracted date
        else:
            raise ValueError("Filename does not contain a valid date.")
            logger.error('Filename does not contain a valid date.')
    except Exception as e:
        print(f"Error extracting timestamp from filename '{filename}': {e}")
        logger.error(filename, '%s Error extracting timestamp from filename %s')
        return datetime.now().replace(hour=7, minute=0, second=0, microsecond=0)

# Folder containing Excel files
excel_folder = "C:/Users/Public/Documents/Reporte Oxidos - Oficial/User Input Folder"
# Folder to move processed Excel files
processed_folder = "C:/Users/Public/Documents/Reporte Oxidos - Oficial/Processed Excel Files"
os.makedirs(processed_folder, exist_ok=True)

# Input parameters file path
input_parameters_file = "Input Parameters.csv"
input_parameters = read_input_parameters("C:/Phyton Scripts/Excel to CSV/bin/Input Parameters.csv")

# Function to convert Excel Ordinal in Date
def from_excel_ordinal(ordinal: float, _epoch0=datetime(1899, 12, 31)) -> datetime:
    if ordinal >= 60:
        ordinal -= 1  # Excel leap year bug, 1900 is not a leap year!
    return (_epoch0 + timedelta(days=ordinal)).replace(microsecond=0)

# Iterate over all Excel files in the folder
def process_excel_files():
    log_date = datetime.now()
    logger = logging.getLogger(__name__)
    FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
    logging.basicConfig(filename=LOG_FILE_PATH, encoding='utf-8', level=logging.DEBUG, format=FORMAT, datefmt='%d/%m/%Y %H:%M:%S')
    logger.debug(f'Script Started')
    
    for excel_filename in os.listdir(excel_folder):
        logger.debug(f'Start of reading Excel File {excel_filename}')
        if excel_filename.endswith('.xlsx'):
            full_file_path = os.path.join(excel_folder, excel_filename)
            # timestamp = extract_timestamp_from_filename(excel_filename)
            try:
                logger.debug(f'Start of reading Excel data using parameters file')
                for param in input_parameters:
                    worksheet = param["Worksheet"]
                    row_index = int(param["Row Index"])
                    column_index = int(param["Column Index"]) - 1
                    output_variable_name = param["Output Variable Name"]

                    # Read the specific worksheet
                    sheet_data = pd.read_excel(full_file_path, sheet_name=worksheet, header=None)
                    
                    timestamp_excel_cell = sheet_data.iloc[3, 6]
                    timestamp = timestamp_excel_cell.replace(hour=7, minute=0, second=0, microsecond=0)

                    # Validate row and column indices
                    if row_index >= len(sheet_data) or column_index >= len(sheet_data.columns):
                        print(
                        f"Error: Row index {row_index + 1} or column index {column_index + 1} "
                        f"out of bounds for file {excel_filename}, sheet {worksheet}."                        
                        )
                        logger.error(f'Row index {row_index + 1} or column index {column_index + 1} out of bounds for file {excel_filename}, sheet {worksheet}.')
                        continue

                    # Extract the specific cell value
                    cell_value = sheet_data.iloc[row_index, column_index]

                    # Prepare timestamp for 7 AM of the current date
                    # timestamp = datetime.now().replace(hour=7, minute=0, second=0, microsecond=0)

                    # Save the extracted value to a .csv file
                    output_data = pd.DataFrame({
                        "Variable Name": [output_variable_name],
                        "Timestamp": [timestamp],
                        "Value": [cell_value]
                    })
                    output_csv_file = f"{os.path.splitext(excel_filename)[0]}.csv"
                    output_csv_file = os.path.join(OUTPUT_FILE_PATH, output_csv_file)
                    # Append to the CSV if it exists, otherwise create it
                    if os.path.exists(output_csv_file):
                        output_data.to_csv(output_csv_file, mode='a', header=False, index=False)
                    else:
                        output_data.to_csv(output_csv_file, index=False)

                    print(f"Processed: {full_file_path} -> {output_csv_file}")
                    logger.debug(f'Processed {full_file_path} -> {output_csv_file}')
            except Exception as e:
                print(f"Error processing file {full_file_path}: {e}")
                logger.error(f'Error processing file {full_file_path}: {e}')
            # Move the processed file to the processed folder
            shutil.move(full_file_path, os.path.join(processed_folder, excel_filename))
            print(f"Moved {full_file_path} to {processed_folder}")        
            logger.debug(f'Moved {full_file_path} to {processed_folder}')
if __name__ == "__main__":
    process_excel_files()