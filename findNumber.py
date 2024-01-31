import os
import pandas as pd

def search_in_folder(folder_path, target_numbers):
    result = []

    # Loop through all files in the folder
    for file_number, filename in enumerate(os.listdir(folder_path), start=1):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(folder_path, filename)

            # Read Excel file using pandas
            df = pd.read_excel(file_path)

            # Search for target numbers in the dataframe
            for target_number in target_numbers:
                locations = [(i, j) for i, row in enumerate(df.values) for j, val in enumerate(row) if val == target_number]
                for row_number, col_number in locations:
                    result.append({
                        'file_number': file_number,
                        'file_name': filename,
                        'line_number': row_number + 2,  # Add 2 to account for 0-based indexing and header row
                        'column_number': col_number + 1,  # Add 1 to account for 0-based indexing
                        'target_number': target_number,
                        'file_path': file_path
                    })

    return result

# Set the folder path and target numbers
folder_path = 'E:/search/'
target_numbers = [3851106, 456, 789]  # Replace with your actual target numbers

# Perform the search
search_result = search_in_folder(folder_path, target_numbers)

# Display the result
for result in search_result:
    print(f"Number {result['target_number']} found in file {result['file_number']}: {result['file_name']} at line {result['line_number']}, column {result['column_number']} in file: {result['file_path']}")
