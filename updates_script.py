import pandas as pd

def generate_summary_tables(input_file, output_file):
    #read the excel file
    df = pd.read_excel(input_file)

    # Ensure "Date" column is in datetime format
    df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%d-%b-%Y")

    ## Convert the other columns to categories except the date column
    categorical_cols = [col for col in df.columns if col != "Date"]

    # Create a writer for excel
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

        for column in categorical_cols:
            # Group by Date and the current column, then count occurrences
            summary_table = df.groupby(["Date", column]).size().unstack(fill_value=0)

            # Add a Total for the row
            summary_table.loc["Total"] = summary_table.sum()

            # Add a Total column
            summary_table["Total"] = summary_table.sum(axis=1)

            # Save each summary table to a new sheet in excel
            summary_table.to_excel(writer, sheet_name=column[:31])
    
    print(f"Summary tables have been saved to {output_file}")

# Add you variables here
input_file = "test.xlsx" #Input your path to excel file here
output_file = "summary_tables.xlsx"
generate_summary_tables(input_file, output_file)

