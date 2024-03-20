import tabula
import pandas as pd

# Path to the PDF file
pdf_path = "/root/saran/408215.pdf"

# Define the range of pages containing the table
pages = "48-64"

# Convert PDF to DataFrame
df = tabula.read_pdf(pdf_path, pages=pages, multiple_tables=True)

# Check if any tables were extracted
if df:
    # Combine all tables into a single DataFrame
    combined_df = pd.concat(df)
    
    # Preprocess DataFrame to handle rowspan and colspan
    # Your logic for handling rowspan and colspan goes here

    # Output Excel file name
    excel_file = "output.xlsx"

    # Write the DataFrame to an Excel file
    combined_df.to_excel(excel_file, index=False)
    print("Data extracted successfully and saved to", excel_file)
else:
    print("No tables found on the specified pages.")

