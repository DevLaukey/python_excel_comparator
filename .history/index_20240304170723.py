import pandas as pd
import matplotlib.pyplot as plt
import openpyxl

def read_excel_file(file_path):
    # if the file is example_overall_results.xlsx then transpose the data
    if "overall" in file_path:
        return pd.read_excel(file_path, index_col=0).T
    else:
        return pd.read_excel(file_path, index_col=0)

def compare_and_visualize(scenario_name, df_operations, df_overall, writer):
    # Your comparison and analysis logic here
    # For example, you can extract specific columns, perform calculations, and plot graphs

    # Extract relevant columns for comparison
    relevant_columns = ["Total revenues (€)", "Day-ahead revenues (€)", "Energy lack balancing revenues (€)",
                        "Energy surplus balancing revenues (€)", "primary up band revenues (€)",
                        "primary down band revenues (€)", "secondary up band revenues (€)",
                        "secondary down band revenues (€)"]

    try:
        df_compare = pd.DataFrame({
            "Operations Results": df_operations.loc[:, relevant_columns].sum(),
            "Overall Results": df_overall.loc[:, relevant_columns].sum()
        })

        # Plotting bar chart for comparison
        plt.figure(figsize=(10, 6))
        df_compare.plot(kind='bar', rot=45, color=['skyblue', 'lightgreen'])
        plt.title(f"Comparison for {scenario_name}")
        plt.ylabel("Revenue (€)")
        plt.savefig(f"Comparison_{scenario_name}.png")
        plt.close()

        # Save the comparison chart to Excel
        sheet_name = f"{scenario_name} Comparison"
        df_compare.to_excel(writer, sheet_name=sheet_name)

    except KeyError as e:
        print(f"\nError: One or more relevant columns not found in the dataframes.")
        print(f"Make sure the specified columns are present in both dataframes.")
        print(f"Error details: {e}")

def create_excel_template(scenario_names):
    # Create a new Excel workbook
    writer = pd.ExcelWriter('Output_Template.xlsx', engine='openpyxl')

    # Write each scenario's comparison to a separate sheet
    for scenario_name in scenario_names:
        df = pd.DataFrame()
        df.to_excel(writer, sheet_name=f"{scenario_name} Comparison")

    # Save the Excel workbook
    writer.save()

def main():
    # Get the number of scenarios to compare
    num_scenarios = int(input("Enter the number of scenarios to compare (1, 2, or 3): "))

    # Get file paths for each scenario
    scenario_files = []
    scenario_names = []
    for i in range(num_scenarios):
        scenario_name = input(f"Enter the name for scenario {i + 1}: ")
        file_path_operations = input(f"Enter the file path for scenario {i + 1} operations results: ")
        file_path_overall = input(f"Enter the file path for scenario {i + 1} overall results: ")
        scenario_files.append((file_path_operations, file_path_overall))
        scenario_names.append(scenario_name)

    # Create the Excel template with empty sheets
    create_excel_template(scenario_names)

    # Open the Excel workbook in append mode
    with pd.ExcelWriter('Output_Template.xlsx', engine='openpyxl', mode='a') as writer:
        for i, (file_path_operations, file_path_overall) in enumerate(scenario_files):
            # Read Excel files for each scenario
            df_operations = read_excel_file(file_path_operations)
            df_overall = read_excel_file(file_path_overall)

            # Compare and visualize results for each scenario
            compare_and_visualize(scenario_names[i], df_operations, df_overall, writer)

if __name__ == "__main__":
    main()
