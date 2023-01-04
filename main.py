import openpyxl

file_name = "HTRs_-_Price_Links_(2).xlsx"

# Open the Excel file
workbook = openpyxl.load_workbook(file_name)

# Get a list of all the sheet names in the workbook
sheet_names = workbook.sheetnames

# Iterate through the sheet names
for sheet_name in sheet_names:
    # Select the sheet
    worksheet = workbook[sheet_name]
    # Iterate through the rows and cells of the sheet
    for row in worksheet.rows:
        for cell in row:
            # Check if the cell value is a hyperlink
            if cell.hyperlink:
                try:
                    # The cell value is a hyperlink, so extract the URL
                    url = cell.hyperlink.target
                    # Replace the hyperlink with the URL
                    cell.value = url
                except:
                    pass

# Save the changes to the Excel file
workbook.save("new_" + file_name)
