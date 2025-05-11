import xlwings as xw
import pandas as pd
from datetime import datetime
import logging
from retry_decorator import retry_xlwings

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Example 1: Decorating a function that gets a range value
@retry_xlwings(tries=3, delay=0.5)
def get_range_value(sheet, range_address):
    """Get the value of a range with retry capability."""
    return sheet.range(range_address).value

# Example 2: Decorating a function that sets a range value
@retry_xlwings(tries=3, delay=0.5)
def set_range_value(sheet, range_address, value):
    """Set the value of a range with retry capability."""
    sheet.range(range_address).value = value
    return True

# Example 3: Decorating a function that clears a sheet
@retry_xlwings(tries=3, delay=0.5)
def clear_sheet(sheet):
    """Clear a sheet with retry capability."""
    sheet.clear()
    return True

# Example 4: Decorating a function that adds a sheet
@retry_xlwings(tries=3, delay=0.5)
def add_sheet(workbook, sheet_name):
    """Add a sheet with retry capability."""
    if sheet_name not in [sheet.name for sheet in workbook.sheets]:
        return workbook.sheets.add(sheet_name)
    return workbook.sheets[sheet_name]

# Example 5: Decorating a function that writes a DataFrame to Excel
@retry_xlwings(tries=3, delay=0.5)
def write_dataframe(sheet, cell, dataframe):
    """Write a DataFrame to Excel with retry capability."""
    sheet.range(cell).value = dataframe
    return True

# Example 6: Creating utility functions for common operations
@retry_xlwings(tries=3, delay=0.5)
def get_sheet(workbook, sheet_name):
    """Get a sheet by name with retry capability."""
    return workbook.sheets[sheet_name]

@retry_xlwings(tries=3, delay=0.5)
def copy_range(source_range, destination_range):
    """Copy a range to another range with retry capability."""
    source_range.copy(destination_range)
    return True

# Example 7: Using the decorator with a more complex function
@retry_xlwings(tries=3, delay=0.5, logger_func=logger.error)
def create_report(workbook, sheet_name, data):
    """Create a report with retry capability."""
    # Check if sheet exists, if not create it
    if sheet_name not in [sheet.name for sheet in workbook.sheets]:
        sheet = workbook.sheets.add(sheet_name)
    else:
        sheet = workbook.sheets[sheet_name]
    
    # Clear the sheet
    sheet.clear()
    
    # Add a title
    sheet.range("A1").value = "Report Generated"
    sheet.range("A2").value = f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    
    # Convert data to DataFrame if it's not already
    if not isinstance(data, pd.DataFrame):
        data = pd.DataFrame(data)
    
    # Write the DataFrame to Excel
    sheet.range("A4").value = data
    
    return sheet

# Example of how to use these functions
def main():
    try:
        # Get the active workbook
        wb = xw.Book.caller()
        
        # Example usage of the decorated functions
        sheet = get_sheet(wb, "Sheet1")
        clear_sheet(sheet)
        set_range_value(sheet, "A1", "Hello World")
        value = get_range_value(sheet, "A1")
        
        print(f"Retrieved value: {value}")
        
        # Create a sample DataFrame
        df = pd.DataFrame({
            'A': [1, 2, 3],
            'B': [4, 5, 6]
        })
        
        # Write the DataFrame
        write_dataframe(sheet, "A3", df)
        
        print("Operations completed successfully")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
