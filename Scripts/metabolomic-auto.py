#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Author: Maxime Naour
# Date(dd-mm-yyyy): 05-10-2023

# Description: Program for processing metabolomics data

# WARNING: Before running the script, make sure to verify the following points:

# 1. Ensure that all necessary input files are placed in the directory specified by the 'path' variable.
# 2. Double-check that the input Excel files containing lipid data are in the correct format and located in the 'path' directory.
# 3. Verify that the 'signal_data' file, which contains signal data, is specified correctly and its format is accurate.
# 4. Confirm that the Excel files have the required columns ('Lipid name', 'Exact Mass' for lipid data, 'Signal name', 'Mass' for signal data).
# 5. Make sure that the input data files do not contain any empty or corrupted cells that might cause errors during processing.
# 6. Ensure that the Python environment has the necessary libraries installed, particularly 'pandas' for data manipulation.
# 7. Validate that the output directory specified by the 'output_path' variable exists and has the necessary write permissions.
# 8. Verify the system compatibility of the file paths, especially if running the script on a different operating system.
# 9. Review the results carefully to ensure the accuracy of the processed data.

# Once you have verified these points, you can proceed to run the script.


# Import the necessary libraries
import os
import glob
import time
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows


# Set the parameters
# Path to the directory containing the files to be processed
path = r"C:\Users\Maxime\OneDrive\Bureau\MetabolomicAuto"

# Change the path to the directory where the files are located
os.chdir(f"{path}")
print("Working directory changed to: ", os.getcwd())

# Path to the directory where the processed files will be saved
# Create a new directory if it doesn't exist already
output_path = "Processed/"
if not os.path.exists(f"{output_path}"):
    os.makedirs(f"{output_path}")

# Sheet name in the Excel files containing the lipid data
lipid_sheet_name = 'Feuil1'

# Sheet name in the Excel files containing the signal data
signal_sheet_name = 'Feuil1'

# Class to read the Excel files in the directory, process them, and save them in a new file in the output directory
class ProcessFiles:
    def __init__(self, path, output_path):
        self.path = path
        self.output_path = output_path
        self.lipid_sheet_name = lipid_sheet_name
        self.signal_sheet_name = signal_sheet_name

    def get_file_names(self):
        self.file_names = glob.glob(f"{self.path}/*.xlsx")
        return self.file_names

    def write_file_name(self):
        print("Available files:")
        print([file.split('\\')[-1] for file in self.file_names])
        self.file_name = input("Enter the file name to process: ")
        return self.file_name

    def read_data_from_excel(self, file_name, sheet_name, columns):
        data = pd.read_excel(file_name, engine='openpyxl', decimal=',', skiprows=4, usecols=columns, sheet_name=sheet_name)
        return data

    def compare_masses(self, lipid_data, signal_data):
        results = []
        for _, signal_row in signal_data.iterrows():
            signal_name = signal_row['Signal name']
            signal_mass = signal_row['Mass']
            confidences = []
            lipid_names = []

            for _, lipid_row in lipid_data.iterrows():
                lipid_name = lipid_row['Lipid name']
                lipid_mass = lipid_row['Exact Mass']
                confidence = (signal_mass - lipid_mass) / signal_mass * 1000000
                if -5 <= confidence <= 5:  # Filter confidences between -5 and 5
                    confidences.append(confidence)
                    lipid_names.append(lipid_name)

            # Sort lipids by confidence, in descending order
            sorted_indices = sorted(range(len(confidences)), key=lambda k: confidences[k], reverse=True)
            sorted_lipid_names = [lipid_names[i] for i in sorted_indices]
            sorted_confidences = [confidences[i] for i in sorted_indices]

            confidence_strings = map(lambda x: str(round(x, 2)), sorted_confidences)
            confidence_string = ' > '.join(confidence_strings)

            results.append([signal_name, signal_mass, ' > '.join(sorted_lipid_names), confidence_string])

        return results

    def save_results(self, results):
        output_df = pd.DataFrame(results, columns=['Signal name', 'Mass', 'Lipid name', 'Confidence'])
        base_name, extension = os.path.splitext(self.file_name)
        output_file = os.path.join(self.output_path, f"Processed_{base_name}.xlsx")

        # Create a new Excel workbook using openpyxl
        workbook = Workbook()

        # Remove the default "Sheet"
        default_sheet = workbook.active
        workbook.remove(default_sheet)

        # Create a new sheet named 'Results'
        sheet = workbook.create_sheet('Results')

        # Write DataFrame data into the Excel sheet
        for record in dataframe_to_rows(output_df, index=False, header=True):
            sheet.append(record)

        # Apply bold style to the first row (column headers)
        for cell in sheet[1]:
            cell.font = Font(bold=True)

        # Access the Excel sheet to adjust column widths and align cells to the left
        for column_cells in sheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = length + 2  # Add space for better spacing
            
            # Align all cells to the left, except the first row
            for i, cell in enumerate(column_cells):
                if cell.row == 1:  # Check if it's the first row
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')  # Center headers
                else:
                    cell.alignment = openpyxl.styles.Alignment(horizontal='left', vertical='center')  # Align other cells to the left

        # Save the Excel workbook
        workbook.save(output_file)
        print(f"Results saved in: {output_file}")

    def run_process(self):
        self.get_file_names()
        file_name = self.write_file_name()
        lipid_data_columns = ['Lipid name', 'Exact Mass']
        signal_data_columns = ['Signal name', 'Mass']
        lipid_data = self.read_data_from_excel(file_name, self.lipid_sheet_name, lipid_data_columns)
        signal_data = self.read_data_from_excel(file_name, self.signal_sheet_name, signal_data_columns)
        results = self.compare_masses(lipid_data, signal_data)
        self.save_results(results)

# Create an instance of the class
process = ProcessFiles(path, output_path)

# Run the process
process.run_process()

# Print the time it took to run the process in seconds
print("Time to run the process: ", time.process_time(), "seconds")
