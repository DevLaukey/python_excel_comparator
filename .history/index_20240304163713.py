import pandas as pd

def read_xlsx_file(file_path):
    return pd.read_excel(file_path)

def get_file_path_from_user():
    file_path = input("Enter the path to the Excel file: ")
    return file_path

if __name__ == "__main__":
    # Get file path from user input
    file_path = get_file_path_from_user()

    # Example usage of read_xlsx_file function
    data = read_xlsx_file(file_path)
    print(data.head())  # print the first 5 rows of the dataframe
