# Import the libraries
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter

def read_input_files(scenario, file_path):
    # Read the overall results file
    overall_results = pd.read_excel(f"{file_path}\\{scenario}_overall_results.xlsx")
    print("Overall Results DataFrame:")
    print(overall_results.head())  # Print first few rows for inspection

    # Read the operations results logs file
    operations_results_logs = pd.read_excel(f"{file_path}\\{scenario}_operations_results_logs.xlsx")
    print("Operations Results Logs DataFrame:")
    print(operations_results_logs.head())  # Print first few rows for inspection

    # Merge the two dataframes on their indices
    merged_df = pd.merge(overall_results, operations_results_logs, left_index=True, right_index=True)
    # Return the merged dataframe
    return merged_df

# Define a function to create a tab for each scenario and compare the results
def create_scenario_tab(scenario, writer, file_path):
    # Call the read_input_files function and get the dataframe
    df = read_input_files(scenario, file_path)
    # Create a new sheet in the output file with the scenario name
    df.to_excel(writer, sheet_name=scenario, index=False)
    # Get the worksheet object
    worksheet = writer.sheets[scenario]
    # Set the column width
    worksheet.set_column("A:Z", 20)
    # Calculate the average, min, max, and std of the price and power outputs
    avg_price = df["Price"].mean()
    min_price = df["Price"].min()
    max_price = df["Price"].max()
    std_price = df["Price"].std()
    avg_power = df["Power"].mean()
    min_power = df["Power"].min()
    max_power = df["Power"].max()
    std_power = df["Power"].std()
    # Calculate the total revenue, cost, and profit
    total_revenue = df["Revenue"].sum()
    total_cost = df["Cost"].sum()
    total_profit = df["Profit"].sum()
    # Calculate the correlation coefficient between the price and power outputs
    corr_price_power = df["Price"].corr(df["Power"])
    # Write the calculations in the worksheet
    worksheet.write("AA1", "Average Price")
    worksheet.write("AA2", avg_price)
    worksheet.write("AB1", "Min Price")
    worksheet.write("AB2", min_price)
    worksheet.write("AC1", "Max Price")
    worksheet.write("AC2", max_price)
    worksheet.write("AD1", "Std Price")
    worksheet.write("AD2", std_price)
    worksheet.write("AE1", "Average Power")
    worksheet.write("AE2", avg_power)
    worksheet.write("AF1", "Min Power")
    worksheet.write("AF2", min_power)
    worksheet.write("AG1", "Max Power")
    worksheet.write("AG2", max_power)
    worksheet.write("AH1", "Std Power")
    worksheet.write("AH2", std_power)
    worksheet.write("AI1", "Total Revenue")
    worksheet.write("AI2", total_revenue)
    worksheet.write("AJ1", "Total Cost")
    worksheet.write("AJ2", total_cost)
    worksheet.write("AK1", "Total Profit")
    worksheet.write("AK2", total_profit)
    worksheet.write("AL1", "Corr Price Power")
    worksheet.write("AL2", corr_price_power)
    # Plot the price and power outputs for the scenario
    fig, ax = plt.subplots()
    ax.plot(df["Time"], df["Price"], label="Price")
    ax.plot(df["Time"], df["Power"], label="Power")
    ax.set_xlabel("Time")
    ax.set_ylabel("Price/Power")
    ax.set_title(f"Price and Power Outputs for {scenario}")
    ax.legend()
    # Save the figure as an image file
    fig.savefig(f"{scenario}_price_power.png")
    # Insert the image file in the worksheet
    worksheet.insert_image("AN1", f"{scenario}_price_power.png")

# Define a function to create an overall-results tab and compare the results across scenarios
def create_overall_results_tab(scenarios, writer, file_path):
    # Create a new sheet in the output file with the name "Overall-Results"
    worksheet = writer.book.add_worksheet("Overall-Results")
    # Set the column width
    worksheet.set_column("A:Z", 20)
    # Define a list of colors for each scenario
    colors = ["green", "blue", "red"]
    # Loop through the scenarios and the colors
    for scenario, color in zip(scenarios, colors):
        # Call the read_input_files function and get the dataframe
        df = read_input_files(scenario, file_path)
        # Write the scenario name in the worksheet
        worksheet.write(f"A{scenarios.index(scenario) + 1}", scenario, writer.book.add_format({"bold": True, "font_color": color}))
        # Write the overall results data in the worksheet
        worksheet.write(f"B{scenarios.index(scenario) + 1}", "Total Revenue")
        worksheet.write(f"C{scenarios.index(scenario) + 1}", df["Revenue"].sum())
        worksheet.write(f"D{scenarios.index(scenario) + 1}", "Total Cost")
        worksheet.write(f"E{scenarios.index(scenario) + 1}", df["Cost"].sum())
        worksheet.write(f"F{scenarios.index(scenario) + 1}", "Total Profit")
        worksheet.write(f"G{scenarios.index(scenario) + 1}", df["Profit"].sum())
        worksheet.write(f"H{scenarios.index(scenario) + 1}", "Average Price")
        worksheet.write(f"I{scenarios.index(scenario) + 1}", df["Price"].mean())
        worksheet.write(f"J{scenarios.index(scenario) + 1}", "Average Power")
        worksheet.write(f"K{scenarios.index(scenario) + 1}", df["Power"].mean())
        worksheet.write(f"L{scenarios.index(scenario) + 1}", "Corr Price Power")
        worksheet.write(f"M{scenarios.index(scenario) + 1}", df["Price"].corr(df["Power"]))
    # Plot the price and power outputs for each scenario
    fig, ax = plt.subplots()
    # Loop through the scenarios and the colors
    for scenario, color in zip(scenarios, colors):
        # Call the read_input_files function and get the dataframe
        df = read_input_files(scenario, file_path)
        # Plot the price and power outputs with the corresponding color and label
        ax.plot(df["Time"], df["Price"], color=color, label=f"{scenario} Price")
        ax.plot(df["Time"], df["Power"], color=color, linestyle="--", label=f"{scenario} Power")
    # Set the xlabel, ylabel, title, and legend for the plot
    ax.set_xlabel("Time")
    ax.set_ylabel("Price/Power")
    ax.set_title("Price and Power Outputs for Each Scenario")
    ax.legend()
    # Save the figure as an image file
    fig.savefig("overall_results_price_power.png")
    # Insert the image file in the worksheet
    worksheet.insert_image("O1", "overall_results_price_power.png")

# Define a function to create the output Excel file
def create_output_file(scenarios, file_path):
    # Create a new Excel file with the name "output.xlsx"
    writer = pd.ExcelWriter("output.xlsx", engine="xlsxwriter")
    # Loop through the scenarios
    for scenario in scenarios:
        # Call the create_scenario_tab function and pass the scenario, the writer, and the file path
        create_scenario_tab(scenario, writer, file_path)
    # Call the create_overall_results_tab function and pass the scenarios, the writer, and the file path
    create_overall_results_tab(scenarios, writer, file_path)
    # Save and close the output file
    writer.save()
    writer.close()

# Ask the user to enter the file path
file_path = input("Enter the file path where the Excel files are located: ")

# Ask the user to choose the number of scenarios to compare and analyze
num_scenarios = int(input("How many scenarios do you want to compare and analyze (1, 2, or 3)? "))
# Check if the number is valid
if num_scenarios not in [1, 2, 3]:
    # If not, print an error message and exit the program
    print("Invalid number. Please enter 1, 2, or 3.")
    exit()

# Create an empty list to store the chosen scenarios
scenarios = []
# Define your scenarios
available_scenarios = ["scenario_1", "scenario_2", "scenario_3"]

# Loop through the range of the number of scenarios
for i in range(num_scenarios):
    # Print the number of the scenario
    print(f"Scenario {i + 1}:")
    # Print the available scenarios for the user to choose from
    print("Available scenarios:")
    for j, scenario in enumerate(available_scenarios):
        print(f"{j + 1}. {scenario}")
    # Ask the user to choose a scenario
    scenario_choice = int(input("Please choose a scenario: "))
    # Check if the choice is valid
    if 1 <= scenario_choice <= len(available_scenarios):
        # Append the chosen scenario to the list
        scenarios.append(available_scenarios[scenario_choice - 1])
    else:
        # If not, print an error message and exit the program
        print("Invalid scenario choice. Please try again.")
        exit()

# Call the create_output_file function and pass the scenarios and the file path
create_output_file(scenarios, file_path)
# Print a success message
print("The output file has been created successfully.")