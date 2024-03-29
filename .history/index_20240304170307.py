import pandas as pd
import matplotlib.pyplot as plt

def read_excel_file(file_path):
    # if the file is example_overall_results.xlsx then transpose the data
    if "overall" in file_path:
        return pd.read_excel(file_path, index_col=0).T
    else:     
        return pd.read_excel(file_path, index_col=0)

def compare_and_visualize(scenario_name, df_operations, df_overall):
    # Print the column names of both dataframes for verification
    print(f"\nColumns in {scenario_name} Operations Results:")
    print(df_operations.columns)

    print(f"\nColumns in {scenario_name} Overall Results:")
    print(df_overall.columns)

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
        df_compare.plot(kind='bar', rot=45)
        plt.title(f"Comparison for {scenario_name}")
        plt.ylabel("Revenue (€)")
        plt.show()

    except KeyError as e:
        print(f"\nError: One or more relevant columns not found in the dataframes.")
        print(f"Make sure the specified columns are present in both dataframes.")
        print(f"Error details: {e}")

        # Print the column names again for reference
        print(f"\nColumns in {scenario_name} Operations Results:")
        print(df_operations.columns)

        print(f"\nColumns in {scenario_name} Overall Results:")
        print(df_overall.columns)


def main():
    # Get the number of scenarios to compare
    num_scenarios = int(input("Enter the number of scenarios to compare (1, 2, or 3): "))

    # Get file paths for each scenario
    scenario_files = []
    for i in range(num_scenarios):
        file_path_operations = input(f"Enter the file path for scenario {i + 1} operations results: ")
        file_path_overall = input(f"Enter the file path for scenario {i + 1} overall results: ")
        scenario_files.append((file_path_operations, file_path_overall))

    for i, (file_path_operations, file_path_overall) in enumerate(scenario_files):
        # Read Excel files for each scenario
        df_operations = read_excel_file(file_path_operations)
        df_overall = read_excel_file(file_path_overall)

        # Compare and visualize results for each scenario
        compare_and_visualize(f"Scenario {i + 1}", df_operations, df_overall)

if __name__ == "__main__":
    main()
