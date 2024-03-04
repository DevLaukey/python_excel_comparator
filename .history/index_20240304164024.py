import pandas as pd

def read_input_files( file_path):
    # Read the overall results file
    overall_results = pd.read_excel(file_path)
    print("Overall Results DataFrame:")
    print(overall_results.head())  # Print first few rows for inspection

def get_file_path_from_user():
    file_path = input("Enter the path to the Excel file: ")
    return file_path

if __name__ == "__main__":
    # Get file path from user input
    file_path = get_file_path_from_user()

    # Example usage of read_xlsx_file function
    data = read_input_files(file_path)
    print(data.head())  # print the first 5 rows of the dataframe
