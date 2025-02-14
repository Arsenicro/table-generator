import sys

from openpyxl import load_workbook


def generate_html_table(file_path, sheet_name):
    # Load the Excel file
    workbook = load_workbook(filename=file_path)

    # Access the specific sheet by name
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        print(f"Sheet '{sheet_name}' not found in the Excel file.")
        sys.exit(1)

    # Generate the header for the HTML table
    html = '<table style="width: 100%; border-collapse: collapse;">\n'
    html += "  <tr>\n"
    html += "    <th></th>\n"
    for i in range(1, 16):
        html += f'    <th style="width: 5%;">{i}</th>\n'
    html += "  </tr>\n"

    # Iterate through each row in the Excel sheet (starting from the second row)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        nr_indeksu = row[2]  # Assuming 'Nr Indeksu' is in the 3rd column (index 2)

        # Skip rows where 'Nr Indeksu' is None or empty
        if not nr_indeksu:
            continue

        html += "  <tr>\n"
        html += f"    <td>{nr_indeksu}</td>\n"

        # Iterate over columns 1 to 15 (adjust indices as needed)
        for col in range(
            6, 21
        ):  # Columns from index 6 to 20 correspond to columns 1-15 in your example
            value = row[col]
            html += f'    <td>{"+" if value else ""}</td>\n'

        html += "  </tr>\n"

    # Close the table tag
    html += "</table>\n"

    return html


if __name__ == "__main__":
    # Check if file path is provided as a command-line argument
    if len(sys.argv) < 2:
        print("Usage: python table_generator.py <excel_file.xlsx>")
        sys.exit(1)

    # Get the file path from command-line arguments
    file_path = sys.argv[1]

    agile = generate_html_table(file_path, "Agile")
    narzedzia = generate_html_table(file_path, "Narzędzia")

    # save agile to agile.txt
    with open("agile.txt", "w") as f:
        f.write(agile)

    # save narzedzia to narzedzia.txt

    with open("narzedzia.txt", "w") as f:
        f.write(narzedzia)
