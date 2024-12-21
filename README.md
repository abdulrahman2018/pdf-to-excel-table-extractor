# PDF to Excel Table Extractor

This project provides a utility to extract tables from PDF files and save them into Excel format. The extractor is designed to work with multi-page PDFs and supports extracting tables from all pages or specific pages.

## Features

- **Extract all tables from a PDF**: The program extracts tables from all pages of a given PDF and saves them in an Excel file.
- **Extract tables from specific pages**: The program can also extract tables from a specified page or a range of pages within the PDF.
- **Clean data**: Empty rows and columns are removed from the extracted tables to ensure the output is cleaner and easier to work with.
- **Handle headers**: The program automatically handles potential header issues, adjusting column headers when necessary.

## Installation

To use this tool, you'll need to have Python and the following libraries installed:

- `tabula-py`: A Python wrapper for Tabula, used to extract tables from PDFs.
- `pandas`: A library for data manipulation, used to handle and write the extracted tables into Excel files.
- `openpyxl`: An Excel file handling library required by Pandas to write data to `.xlsx` files.

Install the required libraries using `pip`:

```bash
pip install tabula-py pandas openpyxl
```

## Usage

To use the extractor, modify the script to provide the path to your PDF file and the desired output Excel file(s). The script provides two main functions:

1. **Extract all tables from the PDF**:
   - This function will extract all the tables from the entire PDF and save them into an Excel file with each table in a separate sheet.

2. **Extract tables from specific pages**:
   - This function allows you to extract tables from specific pages of the PDF and save them into separate Excel files.

## Example

After setting up the required libraries, the script can be executed to extract tables from a given PDF. Here’s an example:

```python
pdf_file = "/path/to/your/pdf.pdf"
output_file_all = "output_all_tables.xlsx"
output_file_page1 = "output_page1_table.xlsx"
output_file_page2 = "output_page2_table.xlsx"
output_file_page3 = "output_page3_table.xlsx"

extract_pdf_to_excel(pdf_file, output_file_all)  # Extract all tables
extract_specific_page(pdf_file, output_file_page1, 1)  # Extract tables from page 1
extract_specific_page(pdf_file, output_file_page2, 2)  # Extract tables from page 2
extract_specific_page(pdf_file, output_file_page3, 3)  # Extract tables from page 3
```

### Expected Output

- Tables from the PDF are saved into Excel files.
- If there are multiple tables on a page, each table is placed in a separate sheet within the Excel file.

## Error Handling

The program handles potential errors such as:

- **Missing PDF file**: If the specified PDF file is not found, an error message will be displayed.
- **No tables found**: If no tables are found in the PDF or the specified page, a message will inform the user.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
