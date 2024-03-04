import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Function to read Excel files and perform basic analysis
def analyze_scenario(scenario_name, overall_results_file, operations_results_file):
    # Read Excel files
    overall_df = pd.read_excel(overall_results_file)
    operations_df = pd.read_excel(operations_results_file)

    # Perform calculations and comparisons
    # Replace the following lines with your specific analysis and calculations
    merged_df = pd.merge(overall_df, operations_df, on='common_column', how='inner')
    calculated_data = merged_df['some_column_from_operations'] - merged_df['some_column_from_overall']

    # Create visualization (example using seaborn)
    sns.scatterplot(x=merged_df['price'], y=calculated_data, hue=scenario_name)
    plt.title(f"Comparison for {scenario_name}")
    plt.show()

    # Return the calculated data for further use
    return calculated_data

# Function to generate the comparison Excel file
def generate_comparison_excel(output_file, scenarios_data):
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for scenario_name, calculated_data in scenarios_data.items():
            calculated_data.to_excel(writer, sheet_name=f"{scenario_name}_comparison", index=False)

# Main function
def main():
    # User input for the number of scenarios to compare
    num_scenarios = int(input("Enter the number of scenarios to compare (1, 2, or 3): "))

    # User input for Excel files
    scenarios_data = {}
    for i in range(num_scenarios):
        scenario_name = input(f"Enter the name for scenario {i + 1}: ")
        overall_file = input(f"Enter the path to overall results file for {scenario_name}: ")
        operations_file = input(f"Enter the path to operations results file for {scenario_name}: ")

        # Perform analysis for each scenario
        calculated_data = analyze_scenario(scenario_name, overall_file, operations_file)
        scenarios_data[scenario_name] = calculated_data

    # Generate the comparison Excel file
    output_excel_file = input("Enter the path for the output Excel file: ")
    generate_comparison_excel(output_excel_file, scenarios_data)

if __name__ == "__main__":
    main()
