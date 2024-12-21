import tabula
import pandas as pd
import os

def extract_pdf_to_excel(pdf_path, output_excel_path):
    """Extracts tables from a PDF and saves them to an Excel file."""
    try:
        # Use tabula to read tables. 'pages="all"' extracts from all pages.
        tables = tabula.read_pdf(pdf_path, pages="all", multiple_tables=True, lattice=True)

        if not tables:
            print(f"No tables found in the PDF: {pdf_path}")
            return

        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                # Clean up empty rows and columns
                table.dropna(axis=0, how='all', inplace=True)
                table.dropna(axis=1, how='all', inplace=True)

                # Handle potential header issues
                if len(table.columns) > 0 and (table.iloc[0].isnull().all() or not isinstance(table.columns[0], str)):
                    table.columns = table.iloc[0].values
                    table = table[1:]
                    table.reset_index(drop=True, inplace=True)

                sheet_name = f"Table_{i+1}" if len(tables) > 1 else "Sheet1"
                table.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Tables extracted and saved to: {output_excel_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

def extract_specific_page(pdf_path, output_excel_path, page_number):
    """Extracts tables from a specific page of a PDF."""
    try:
        tables = tabula.read_pdf(pdf_path, pages=page_number, multiple_tables=True, lattice=True)

        if not tables:
            print(f"No tables found on page {page_number} of {pdf_path}")
            return

        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                # Clean up empty rows and columns
                table.dropna(axis=0, how='all', inplace=True)
                table.dropna(axis=1, how='all', inplace=True)

                 # Handle potential header issues
                if len(table.columns) > 0 and (table.iloc[0].isnull().all() or not isinstance(table.columns[0], str)):
                    table.columns = table.iloc[0].values
                    table = table[1:]
                    table.reset_index(drop=True, inplace=True)

                sheet_name = f"Table_{i+1}" if len(tables) > 1 else "Sheet1" #if there is more than one table on the page it will create different sheets
                table.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Tables from page {page_number} extracted and saved to: {output_excel_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    pdf_file = "/Users/abdulrahmankhaled/my_new_project/myenv/Attendance Sheet _ Strategy.pdf"
    output_file_all = "output_tables_all.xlsx"
    output_file_page1 = "output_table_page1.xlsx"
    output_file_page2 = "output_table_page2.xlsx"
    output_file_page3 = "output_table_page3.xlsx"


    if not os.path.exists(pdf_file):
        print(f"Error: PDF file not found at: {pdf_file}")
    else:
        # Extract all pages
        extract_pdf_to_excel(pdf_file, output_file_all)

        # Extract specific pages
        extract_specific_page(pdf_file, output_file_page1, 1)  # Page 1
        extract_specific_page(pdf_file, output_file_page2, 2)  # Page 2
        extract_specific_page(pdf_file, output_file_page3, 3)  # Page 3
