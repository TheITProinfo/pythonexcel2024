import os
import xlwings as xw

def copy_sheets(source_file, target_file):
    # Get the full file paths
    source_file = os.path.abspath(source_file)
    target_file = os.path.abspath(target_file)
    
    # Open the source workbook
    wb_source = xw.Book(source_file)
    
    # Open the target workbook
    wb_target = xw.Book(target_file)
    
    # Iterate through all sheets in the source workbook
    for sheet in wb_source.sheets:
        # Add a new sheet in the target workbook with the same name
        wb_target.sheets.add(name=sheet.name)
        
        # Get the new sheet in the target workbook
        new_sheet = wb_target.sheets[sheet.name]
        
        # Copy values, formulas, and formats from source to target sheet
        new_sheet.range("A1").value = sheet.range("A1").expand().value
        
        # Copy other properties if needed, e.g., column widths
        new_sheet.api.Columns.ColumnWidth = sheet.api.Columns.ColumnWidth
    
    # Save the target workbook
    wb_target.save()
    
    # Close the workbooks
    wb_target.close()
    wb_source.close()

if __name__ == "__main__":
    source_file = "salesinfo.xlsx"
    target_file = "testresult.xlsx"
    
    # Get the current working directory
    current_dir = os.getcwd()
    current_dir=os.path.join(current_dir,'salesinfo')
    print(current_dir)
    
    # Full paths to source and target files
    source_file = os.path.join(current_dir, source_file)
    target_file = os.path.join(current_dir, target_file)
    
    copy_sheets(source_file, target_file)
    print("Sheets copied successfully!")
