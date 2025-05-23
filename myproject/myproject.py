
import xlwings as xw
import pandas as pd
import random
import tkinter as tk
from tkinter import scrolledtext
from datetime import datetime, timedelta
import traceback
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
import pyperclip
import logging
from retry_decorator import retry_xlwings

# Setup logging for retry decorator
# Configure basic logging to console
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        # Stream handler - logs to console
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# Function to keep the UI responsive during long-running operations
def keep_ui_responsive(func, status_label, btn_text, root):
    # Create a flag to track completion
    is_complete = threading.Event()
    result = [None]  # Use a list to store the result (to allow modification from inner function)
    error = [None]   # Use a list to store any error

    # Function to run in a separate thread
    def threaded_task():
        try:
            result[0] = func()
        except Exception as e:
            error[0] = e
        finally:
            is_complete.set()

    # Start the thread
    thread = threading.Thread(target=threaded_task)
    thread.daemon = True
    thread.start()

    # Update progress while waiting for completion
    start_time = time.time()
    last_update_time = 0
    while not is_complete.is_set():
        # Get elapsed time
        elapsed = time.time() - start_time

        # Update the status with elapsed time (mm:ss format)
        minutes = int(elapsed) // 60
        seconds = int(elapsed) % 60
        status_text = f"Processing time: {minutes:02d}:{seconds:02d}"
        status_label.config(text=status_text)

        # Process any pending UI events to keep the window responsive
        root.update()

        # Sleep briefly to avoid consuming too much CPU
        time.sleep(0.1)

        # Every 2 seconds, print a progress message to console
        current_second = int(elapsed)
        if current_second != last_update_time and current_second % 2 == 0:
            print(f"Still processing... ({minutes:02d}:{seconds:02d} elapsed)")
            last_update_time = current_second

    # Check if there was an error
    if error[0] is not None:
        raise error[0]

    # Update status to "Completed"
    status_label.config(text=f"Completed: {btn_text}")

    return result[0]


def func1():
    """Creates a random financial dataframe with stock prices and saves it to Excel."""
    # Add 50% chance of generating an exception for testing
    if random.random() < 0.5:
        raise Exception("Test exception in func1: This is a simulated error to test exception handling")

    # Create a random dataframe with financial data
    num_days = 30
    today = datetime.now()
    dates = [(today - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(num_days)]

    stocks = ['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'META']
    data = {}

    for stock in stocks:
        # Generate random stock prices with some trend
        base_price = random.uniform(100, 500)
        prices = []
        for _ in range(num_days):
            change = random.uniform(-5, 5)
            base_price += change
            prices.append(round(base_price, 2))
        data[stock] = prices

    # Create DataFrame
    df = pd.DataFrame(data, index=dates)

    # Save to Excel
    wb = xw.Book.caller()
    sheet_name = "Stock_Prices"

    # Check if sheet exists, if not create it
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

    sheet = wb.sheets[sheet_name]
    sheet.clear()
    sheet.range("A1").value = "Date Generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.range("A3").value = df


def func2():
    """Creates a random financial dataframe with portfolio performance and saves it to Excel."""
    # Add 50% chance of generating an exception for testing
    if random.random() < 0.5:
        raise Exception("Test exception in func2: This is a simulated error to test exception handling")

    # Create a random dataframe with portfolio performance data
    num_months = 12
    today = datetime.now()
    months = [(today.replace(day=1) - timedelta(days=30*i)).strftime('%Y-%m') for i in range(num_months)]

    portfolios = ['Conservative', 'Balanced', 'Growth', 'Aggressive']
    data = {}

    for portfolio in portfolios:
        # Generate random returns with different volatility based on portfolio type
        volatility = {
            'Conservative': 2,
            'Balanced': 5,
            'Growth': 8,
            'Aggressive': 12
        }[portfolio]

        returns = [round(random.uniform(-volatility, volatility+2), 2) for _ in range(num_months)]
        data[portfolio] = returns

    # Create DataFrame
    df = pd.DataFrame(data, index=months)
    df.index.name = 'Month'

    # Add cumulative returns
    cumulative_df = pd.DataFrame()
    for portfolio in portfolios:
        cumulative_returns = [100]
        for ret in df[portfolio]:
            cumulative_returns.append(cumulative_returns[-1] * (1 + ret/100))
        cumulative_df[portfolio] = cumulative_returns[1:]

    cumulative_df.index = months

    # Save to Excel
    wb = xw.Book.caller()
    sheet_name = "Portfolio_Performance"

    # Check if sheet exists, if not create it
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

    sheet = wb.sheets[sheet_name]
    sheet.clear()
    sheet.range("A1").value = "Date Generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.range("A3").value = "Monthly Returns (%)"
    sheet.range("A4").value = df

    sheet.range("A" + str(7 + len(df))).value = "Cumulative Performance (Starting at 100)"
    sheet.range("A" + str(8 + len(df))).value = cumulative_df


def func3():
    """Dummy function for correlation analysis."""
    # Create a simple message in Excel
    wb = xw.Book.caller()
    sheet_name = "Correlation_Analysis"

    # Check if sheet exists, if not create it - using retry-enabled functions
    try:
        # Try to get the sheet first
        sheet = get_sheet(wb, sheet_name)
    except Exception:
        # If it doesn't exist, add it
        sheet = add_sheet(wb, sheet_name)

    # Clear the sheet with retry capability
    clear_sheet(sheet)

    # Set values with retry capability
    set_range_value(sheet, "A1", "Correlation Analysis Function Called")
    set_range_value(sheet, "A2", "Date: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))


def func4():
    """Dummy function for risk metrics."""
    # Create a simple message in Excel
    wb = xw.Book.caller()
    sheet_name = "Risk_Metrics"

    # Check if sheet exists, if not create it
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

    sheet = wb.sheets[sheet_name]
    sheet.clear()
    sheet.range("A1").value = "Risk Metrics Function Called"
    sheet.range("A2").value = "Date: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def func5():
    """Dummy function for scenario analysis that deletes and recreates the sheet."""
    # Get the Excel workbook
    wb = xw.Book.caller()
    sheet_name = "Scenario_Analysis"

    # Delete the sheet if it exists (warnings are suppressed at the button handler level)
    if sheet_name in [sheet.name for sheet in wb.sheets]:
        wb.sheets[sheet_name].delete()
        print(f"Deleted sheet: {sheet_name}")

    # Create a new sheet
    new_sheet = wb.sheets.add(sheet_name)
    print(f"Created new sheet: {sheet_name}")

    # Add content to the sheet
    new_sheet.range("A1").value = "Scenario Analysis Function Called"
    new_sheet.range("A2").value = "Date: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_sheet.range("A3").value = "This sheet was deleted and recreated without warnings"


def func6():
    """Dummy function for optimization."""
    # Create a simple message in Excel
    wb = xw.Book.caller()
    sheet_name = "Optimization"

    # Check if sheet exists, if not create it
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

    sheet = wb.sheets[sheet_name]
    sheet.clear()
    sheet.range("A1").value = "Optimization Function Called"
    sheet.range("A2").value = "Date: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def func7():
    """Dummy function for reporting that uses threading for file operations."""
    # Create a simple message in Excel
    wb = xw.Book.caller()
    sheet_name = "Reporting"

    # Check if sheet exists, if not create it
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

    sheet = wb.sheets[sheet_name]
    sheet.clear()
    sheet.range("A1").value = "Threaded Reporting Function Called"
    sheet.range("A2").value = "Date: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Simulate reading multiple files with ThreadPoolExecutor
    def read_file(file_num):
        # Simulate file reading by sleeping
        time.sleep(1)  # Simulate a 1-second file read
        return f"Data from file {file_num}"

    # Use ThreadPoolExecutor to read files in parallel
    results = []
    with ThreadPoolExecutor(max_workers=4) as executor:
        # Submit 10 file reading tasks
        futures = [executor.submit(read_file, i) for i in range(1, 11)]

        # Collect results as they complete
        for i, future in enumerate(futures):
            result = future.result()
            results.append(result)
            print(f"Read file {i+1}/10: {result}")

    # Write results to Excel
    for i, result in enumerate(results):
        sheet.range(f"A{i+4}").value = result

    sheet.range("A3").value = f"Read {len(results)} files successfully"


@xw.func
def hello(name):
    return f"Hello {name}!"


@retry_xlwings(tries=4, delay=0.5, backoff=1.5, jitter=0.2, logger_func=logger.warning)
def func8():
    """Example function with direct retry decorator application."""
    # Create a simple message in Excel
    wb = xw.Book.caller()
    sheet_name = "Retry_Example"

    # Check if sheet exists, if not create it
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

    sheet = wb.sheets[sheet_name]
    sheet.clear()

    # This entire function will be retried if any COM errors occur
    sheet.range("A1").value = "Retry Decorator Example"
    sheet.range("A2").value = "Date: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Create a sample DataFrame
    df = pd.DataFrame({
        'Column1': [1, 2, 3, 4, 5],
        'Column2': [10, 20, 30, 40, 50],
        'Column3': [100, 200, 300, 400, 500]
    })

    # Write the DataFrame to Excel
    sheet.range("A-100").value = df

    # Format headers
    sheet.range("A4:C4").api.Font.Bold = True

    # Add a note about retry functionality
    sheet.range("A10").value = "This function demonstrates the retry decorator."
    sheet.range("A10").api.Font.Bold = True
    sheet.range("A11").value = "Check the logs folder for error logs with retry information."

    return "Function completed with retry capability"


def log_exception(exception_text, workbook_path, exc_value=None):
    """Log exception to a file in a logs subfolder."""
    try:
        # Get the directory of the workbook
        workbook_dir = os.path.dirname(os.path.abspath(workbook_path))

        # Create logs subfolder if it doesn't exist
        logs_dir = os.path.join(workbook_dir, "logs")
        os.makedirs(logs_dir, exist_ok=True)

        # Create a timestamp for the filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        log_filename = os.path.join(logs_dir, f"error_log_{timestamp}.txt")

        # Write the exception to the log file
        with open(log_filename, 'w') as log_file:
            log_file.write(f"Exception occurred at {datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')}\n")
            log_file.write("=" * 80 + "\n")
            log_file.write(exception_text)

            # Check if there are retry logs attached to the exception
            if exc_value is not None and hasattr(exc_value, 'retry_logs') and exc_value.retry_logs:
                log_file.write("\n\n" + "=" * 80 + "\n")
                log_file.write("RETRY ATTEMPTS:\n")
                log_file.write("=" * 80 + "\n")
                for i, retry_log in enumerate(exc_value.retry_logs):
                    log_file.write(f"\nRetry Attempt #{i+1}:\n")
                    log_file.write(retry_log)
                    log_file.write("\n")

            log_file.write("\n" + "=" * 80)

        return log_filename
    except Exception as e:
        # If logging fails, print to console as a fallback
        print(f"Failed to log exception: {str(e)}")
        print(exception_text)
        return None


def show_exception_dialog(exception_text, log_filename, exc_value=None):
    """Show a modal dialog with the exception details."""
    # Create a new top-level window
    dialog = tk.Toplevel()
    dialog.title("Error Occurred")
    dialog.geometry("700x500")  # Larger size to accommodate retry logs
    dialog.minsize(500, 400)
    dialog.grab_set()  # Make the dialog modal
    dialog.focus_set()

    # Create a frame for the content
    frame = tk.Frame(dialog, padx=10, pady=10)
    frame.pack(expand=True, fill="both")

    # Add a label explaining the error
    error_label = tk.Label(
        frame,
        text="An error occurred while executing the function.",
        font=("Arial", 12, "bold"),
        fg="red"
    )
    error_label.pack(pady=(0, 10))

    # Add information about the log file
    if log_filename:
        log_label = tk.Label(
            frame,
            text=f"The full error has been logged to:\n{log_filename}",
            font=("Arial", 10),
            justify="left"
        )
        log_label.pack(pady=(0, 10))

    # Create a scrolled text widget to display the exception
    text_area = scrolledtext.ScrolledText(
        frame,
        wrap=tk.WORD,
        width=80,
        height=20,
        font=("Courier New", 10)
    )
    text_area.pack(expand=True, fill="both")

    # Insert the exception text
    text_area.insert(tk.END, exception_text)

    # If there are retry logs, add them to the text area
    if exc_value is not None and hasattr(exc_value, 'retry_logs') and exc_value.retry_logs:
        text_area.insert(tk.END, "\n\n" + "=" * 80 + "\n")
        text_area.insert(tk.END, "RETRY ATTEMPTS:\n")
        text_area.insert(tk.END, "=" * 80 + "\n\n")

        for i, retry_log in enumerate(exc_value.retry_logs):
            text_area.insert(tk.END, f"Retry Attempt #{i+1}:\n")
            text_area.insert(tk.END, retry_log)
            text_area.insert(tk.END, "\n\n")

        text_area.insert(tk.END, "=" * 80)

    text_area.config(state="disabled")  # Make it read-only

    # Scroll to the top
    text_area.see("1.0")

    # Create a frame for buttons
    button_frame = tk.Frame(frame)
    button_frame.pack(pady=10)

    # Copy button
    def copy_to_clipboard():
        # Get all text from the text area
        text_area.config(state="normal")
        full_text = text_area.get("1.0", tk.END)
        text_area.config(state="disabled")

        pyperclip.copy(full_text)
        copy_btn.config(text="Copied!")
        dialog.after(1000, lambda: copy_btn.config(text="Copy to Clipboard"))

    copy_btn = tk.Button(
        button_frame,
        text="Copy to Clipboard",
        command=copy_to_clipboard,
        width=15
    )
    copy_btn.pack(side="left", padx=5)

    # Close button
    close_btn = tk.Button(
        button_frame,
        text="Close",
        command=dialog.destroy,
        width=15
    )
    close_btn.pack(side="left", padx=5)


# Function to check if Excel is still open
def is_excel_still_open(wb_path):
    try:
        # Try to get all open Excel workbooks
        all_books = xw.books
        # Check if our workbook is still open
        for book in all_books:
            if hasattr(book, 'fullname') and book.fullname.lower() == wb_path.lower():
                return True
        return False
    except Exception as e:
        print(f"Error checking Excel: {e}")
        return False


# Utility functions with retry for common xlwings operations
@retry_xlwings(tries=3, delay=0.5, logger_func=logger.warning)
def get_sheet(workbook, sheet_name):
    """Get a sheet by name with retry capability."""
    return workbook.sheets[sheet_name]


@retry_xlwings(tries=3, delay=0.5, logger_func=logger.warning)
def add_sheet(workbook, sheet_name):
    """Add a sheet with retry capability."""
    return workbook.sheets.add(sheet_name)


@retry_xlwings(tries=3, delay=0.5, logger_func=logger.warning)
def clear_sheet(sheet):
    """Clear a sheet with retry capability."""
    sheet.clear()
    return True


@retry_xlwings(tries=3, delay=0.5, logger_func=logger.warning)
def set_range_value(sheet, range_address, value):
    """Set the value of a range with retry capability."""
    sheet.range(range_address).value = value
    return True


@retry_xlwings(tries=3, delay=0.5, logger_func=logger.warning)
def get_range_value(sheet, range_address):
    """Get the value of a range with retry capability."""
    return sheet.range(range_address).value

def main():
    """
    Main function to create the Excel ribbon UI.

    The function checks command-line arguments for a --tabs flag:
    --tabs=Market~Portfolio will only show the Market and Portfolio tabs.
    If the --tabs flag is not present or empty, all tabs will be shown.
    """
    # Check for command-line arguments with --tabs flag
    enabled_tabs = None

    # Parse all command-line arguments looking for --tabs flag
    for arg in sys.argv:
        if arg.startswith("--tabs="):
            # Extract the value after --tabs=
            tabs_value = arg.split("=", 1)[1].strip()
            if tabs_value:  # Only set if not empty
                enabled_tabs = tabs_value
                print(f"Found --tabs flag with value: {enabled_tabs}")
            break
    # Define a list of preset colors for buttons (background, foreground)
    button_colors = [
        ("#4F81BD", "white"),  # Blue
        ("#C0504D", "white"),  # Red
        ("#9BBB59", "black"),  # Green
        ("#8064A2", "white"),  # Purple
        ("#4BACC6", "black"),  # Turquoise
        ("#F79646", "black"),  # Orange
        ("#FFFF00", "black"),  # Yellow
        ("#C00000", "white"),  # Dark Red
        ("#0070C0", "white"),  # Dark Blue
        ("#00B050", "black"),  # Dark Green
    ]

    # Get the Excel workbook that called this function
    try:
        wb = xw.Book.caller()
        # Get the full path of the workbook for display
        wb_path = wb.fullname
        print(f"Workbook path: {wb_path}")
    except Exception as e:
        print(f"Error getting caller workbook: {e}")
        wb_path = "Unknown Workbook"

    # Parse the enabled_tabs parameter
    enabled_tab_list = []
    if enabled_tabs:
        # Split the tilde-delimited string into a list
        enabled_tab_list = [tab.strip() for tab in enabled_tabs.split('~') if tab.strip()]
        print(f"Enabled tabs: {enabled_tab_list}")

    # Define the function map with tabs as keys and button definitions as values
    # Format: {"Tab Name": {"Button Label": function_reference, ...}, ...}
    funcs_map = {
        "Market": {
            "Stocks": func1,
            "Correl": func3,
            "Risk": func4,
        },
        "Portfolio": {
            "Perf": func2,
            "Scenario": func5,
            "Optim": func6,
            "Report": func7,
        },
        "Utils": {
            "Retry": func8,
        }
    }

    # Get the Excel file name for the title
    excel_file_name = os.path.basename(wb_path)

    # Create a simple GUI with tkinter
    root = tk.Tk()
    root.title(f"{excel_file_name}")

    # Initially create with a minimal size - we'll resize it later
    root.geometry("100x100")  # Start with a small window
    root.minsize(100, 100)    # Set minimum size

    # Set a light gray background to mimic Excel's ribbon
    root.configure(bg="#f0f0f0")

    # Position window at top-left corner (0,0)
    root.geometry("+0+0")

    # Make window stay on top
    root.attributes("-topmost", True)

    # No need to periodically check if Excel is closed
    # The window will close when Excel closes due to the COM connection

    # Handle window close event
    def on_window_close():
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_window_close)

    # Create main frame with Excel ribbon style
    main_frame = tk.Frame(root, padx=5, pady=5, bg="#f0f0f0")
    main_frame.pack(expand=True, fill="both")

    # Create a label with the Excel file name
    header_label = tk.Label(main_frame, text=f"{excel_file_name}", font=("Arial", 12, "bold"), bg="#f0f0f0")
    header_label.pack(pady=5)

    # Create a custom notebook with colored tabs
    class ColoredNotebook(tk.Frame):
        def __init__(self, parent, tab_colors, **kwargs):
            tk.Frame.__init__(self, parent, **kwargs)
            self.tab_colors = tab_colors

            # Create a centered frame for the tabs
            self.outer_tab_frame = tk.Frame(self, bg="#f0f0f0")
            self.outer_tab_frame.pack(fill="x", side="top")

            # Create a centered inner frame for the tabs
            self.tab_frame = tk.Frame(self.outer_tab_frame, bg="#f0f0f0")
            self.tab_frame.pack(side="top")

            # Create a frame for the content
            self.content_frame = tk.Frame(self, bg="#f0f0f0")
            self.content_frame.pack(fill="both", expand=True)

            self.tabs = []
            self.tab_buttons = []
            self.current_tab = None

        def add(self, frame, text, tab_color, text_color):
            # Create a tab button with the specified color and rounded corners
            tab_index = len(self.tabs)

            # Create a canvas for the rounded button
            canvas_width = 120  # Width of the button
            canvas_height = 50  # Height of the button
            corner_radius = 10  # Radius of the rounded corners

            canvas = tk.Canvas(self.tab_frame, width=canvas_width, height=canvas_height,
                             bg="#f0f0f0", highlightthickness=0)
            canvas.pack(side="left", padx=2, pady=2)

            # Create rounded rectangle on canvas
            canvas.create_rectangle(
                corner_radius, 0,
                canvas_width - corner_radius, canvas_height,
                fill=tab_color, outline=""
            )
            canvas.create_rectangle(
                0, corner_radius,
                canvas_width, canvas_height - corner_radius,
                fill=tab_color, outline=""
            )

            # Create rounded corners
            canvas.create_arc(
                0, 0,
                corner_radius * 2, corner_radius * 2,
                start=90, extent=90, fill=tab_color, outline=""
            )
            canvas.create_arc(
                canvas_width - corner_radius * 2, 0,
                canvas_width, corner_radius * 2,
                start=0, extent=90, fill=tab_color, outline=""
            )
            canvas.create_arc(
                0, canvas_height - corner_radius * 2,
                corner_radius * 2, canvas_height,
                start=180, extent=90, fill=tab_color, outline=""
            )
            canvas.create_arc(
                canvas_width - corner_radius * 2, canvas_height - corner_radius * 2,
                canvas_width, canvas_height,
                start=270, extent=90, fill=tab_color, outline=""
            )

            # Add text to the canvas
            canvas.create_text(
                canvas_width // 2, canvas_height // 2,
                text=text, fill=text_color,
                font=("Arial", 9, "bold")
            )

            # Store the canvas as the tab button
            tab_button = canvas

            # Bind click event to the canvas
            # Using _ as a prefix for event indicates it's intentionally unused
            canvas.bind("<Button-1>", lambda _, idx=tab_index: self.select_tab(idx))

            # After all tabs are added, we'll center them

            # Hide the frame initially
            frame.pack_forget()

            # Store the tab information
            self.tabs.append(frame)
            self.tab_buttons.append(tab_button)

            # If this is the first tab, select it
            if tab_index == 0:
                self.select_tab(0)

            return tab_index

        def select_tab(self, index):
            # Hide the current tab if there is one
            if self.current_tab is not None:
                self.tabs[self.current_tab].pack_forget()

                # Reset the previous tab's appearance
                prev_canvas = self.tab_buttons[self.current_tab]
                prev_canvas.delete("highlight")

            # Show the selected tab
            self.tabs[index].pack(fill="both", expand=True, in_=self.content_frame)

            # Highlight the selected tab
            selected_canvas = self.tab_buttons[index]
            selected_canvas.create_rectangle(
                2, 2, selected_canvas.winfo_width()-2, selected_canvas.winfo_height()-2,
                outline="white", width=2, tags="highlight"
            )

            # Get the color of the selected tab for the divider line
            # For canvas, we need to get the fill color of the rectangle
            selected_color = self.tab_colors[index][0]  # Use the background color from the color tuple

            # Add or update divider line
            if hasattr(self, 'divider_line'):
                self.divider_line.config(bg=selected_color)
            else:
                self.divider_line = tk.Frame(self, height=2, bg=selected_color)
                self.divider_line.pack(fill="x", side="top", before=self.content_frame)

            self.current_tab = index

    # Create the notebook with our custom class
    notebook = ColoredNotebook(main_frame, button_colors, bg="#f0f0f0")
    notebook.pack(expand=True, fill="both", padx=5, pady=5)

    # After all tabs are created, center them
    notebook.tab_frame.pack_forget()
    notebook.tab_frame.pack(side="top", anchor="center")

    # Status label at the bottom - will be set later
    status_label = None

    # Function to handle button clicks with status updates
    def create_button_handler(func, btn_text):
        def button_handler():
            # Disable all buttons to prevent user interaction during processing
            for tab_buttons in all_buttons:
                for btn in tab_buttons:
                    btn.config(state="disabled")

            # Update status to "Processing..."
            status_label.config(text="Processing...")
            root.update()

            # Check if Excel is still open before proceeding
            if not is_excel_still_open(wb_path):
                # Excel is closed, close the window
                print(f"Excel closed, destroying window during button click")
                root.destroy()
                return

            # Get the Excel application and disable user interaction
            excel_app = None
            try:
                excel_app = xw.apps.active
                if excel_app:
                    excel_app.interactive = False
            except Exception as e:
                print(f"Error getting Excel app: {e}")  # If we can't get the Excel app, continue anyway

            # Define a function to execute the Excel operation on the main thread
            def do_excel_operation():
                nonlocal excel_app

                try:
                    # Check again if Excel is still open
                    if not is_excel_still_open(wb_path):
                        # Excel is closed, close the window
                        print(f"Excel closed, destroying window during function execution")
                        root.destroy()
                        return

                    # Call the function with Excel warnings suppressed
                    try:
                        # Store the original Excel settings
                        original_alerts = None
                        original_enable_events = None
                        try:
                            # Get the Excel application and disable alerts and events
                            if excel_app:
                                # Disable alerts
                                original_alerts = excel_app.display_alerts
                                excel_app.display_alerts = False
                                print(f"Disabled Excel display alerts")

                                # Disable events
                                original_enable_events = excel_app.enable_events
                                excel_app.enable_events = False
                                print(f"Disabled Excel events")
                        except Exception as e:
                            print(f"Error disabling Excel alerts/events: {e}")

                        # For all functions, use the keep_ui_responsive approach to ensure
                        # the UI stays responsive during execution
                        keep_ui_responsive(func, status_label, btn_text, root)
                    except Exception as e:
                        print(f"Exception in function: {e}")
                        # Get the full exception traceback
                        exc_type, exc_value, exc_traceback = sys.exc_info()
                        exception_text = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))

                        # Log the exception to a file using only the caller workbook path
                        # This is the only acceptable approach in this context
                        log_path = wb.fullname

                        # Pass the exception object to log_exception to include retry logs
                        log_filename = log_exception(exception_text, log_path, exc_value)

                        # Update status label
                        status_label.config(text=f"Error occurred. See log for details.")

                        # Show the exception dialog with retry logs
                        show_exception_dialog(exception_text, log_filename, exc_value)
                finally:
                    # Re-enable Excel user interaction and alerts
                    try:
                        if excel_app:
                            # Restore interactive mode
                            excel_app.interactive = True

                            # Restore display alerts if we changed it
                            if original_alerts is not None:
                                excel_app.display_alerts = original_alerts
                                print(f"Restored Excel display alerts to: {original_alerts}")

                            # Restore events if we changed it
                            if original_enable_events is not None:
                                excel_app.enable_events = original_enable_events
                                print(f"Restored Excel events to: {original_enable_events}")
                    except Exception as e:
                        print(f"Error restoring Excel settings: {e}")  # If we can't set the Excel app, continue anyway

                    # Re-enable all buttons
                    for tab_buttons in all_buttons:
                        for btn in tab_buttons:
                            btn.config(state="normal")
                    print("Buttons re-enabled")

                    # Schedule reset of status label after 3 seconds if it's showing completed
                    if status_label.cget("text").startswith("Completed"):
                        root.after(3000, lambda: status_label.config(text="Ready"))

            # Disable all buttons to prevent user interaction during processing
            for tab_buttons in all_buttons:
                for btn in tab_buttons:
                    btn.config(state="disabled")

            # Execute the Excel operation directly on the main thread
            # This avoids COM threading issues
            do_excel_operation()

        return button_handler

    # Button colors already defined at the beginning of the function

    # Store all buttons for later access (to enable/disable them)
    all_buttons = []

    # Create tabs and buttons
    tab_count = 0
    for tab_name, button_dict in funcs_map.items():
        # Skip this tab if it's not in the enabled_tab_list (when the list is not empty)
        if enabled_tab_list and tab_name not in enabled_tab_list:
            print(f"Skipping tab: {tab_name} (not in enabled tabs list)")
            continue

        # Get color for this tab (cycle through the preset colors)
        bg_color, fg_color = button_colors[tab_count % len(button_colors)]
        tab_count += 1

        # Create a tab frame
        tab_frame = tk.Frame(notebook, bg="#f0f0f0")
        notebook.add(tab_frame, text=tab_name, tab_color=bg_color, text_color=fg_color)

        # Create a frame for buttons in this tab with less padding
        buttons_frame = tk.Frame(tab_frame, padx=5, pady=5)
        buttons_frame.pack(expand=True, fill="both")

        # Create buttons in a grid layout
        col = 0
        max_cols = 10  # Maximum number of buttons per row

        # Calculate how many buttons we have for this tab
        num_buttons = len(button_dict)

        # Calculate how many columns to use for this row (for centering)
        cols_this_row = min(num_buttons, max_cols)

        # Calculate starting column for centering buttons
        start_col = (max_cols - cols_this_row) // 2 if cols_this_row < max_cols else 0

        # Store buttons for this tab
        tab_buttons = []

        # Button counter for color cycling
        btn_count = 0

        # Create a frame for buttons in this tab
        buttons_frame = tk.Frame(tab_frame, padx=10, pady=10)
        buttons_frame.pack(expand=True, fill="both")

        for btn_text, btn_func in button_dict.items():
            # Get color for this button (cycle through the preset colors)
            bg_color, fg_color = button_colors[btn_count % len(button_colors)]
            btn_count += 1

            # Create button with Excel ribbon style and color
            btn = tk.Button(
                buttons_frame,
                text=btn_text,
                command=create_button_handler(btn_func, btn_text),
                width=10,  # Fixed width for all buttons
                height=2,  # Fixed height
                font=("Arial", 8, "bold"),
                relief="raised",  # Raised appearance
                borderwidth=1,    # Visible border
                bg=bg_color,      # Background color from preset
                fg=fg_color,      # Text color from preset
                activebackground=bg_color,
                activeforeground=fg_color
            )
            # If we're centering, adjust the column
            grid_col = col
            if cols_this_row < max_cols:
                grid_col = start_col + col

            btn.grid(row=0, column=grid_col, padx=5, pady=5, sticky="nsew")
            tab_buttons.append(btn)

            # Update column for next button
            col += 1

        all_buttons.append(tab_buttons)

        # Configure grid weights to make buttons expand properly
        for i in range(max_cols):
            buttons_frame.columnconfigure(i, weight=1)

    # Create a single status label at the bottom with minimal padding
    status_label = tk.Label(main_frame, text="Ready", font=("Arial", 10))
    status_label.pack(pady=0)  # No padding

    # Calculate the maximum number of buttons in any tab
    max_buttons = max(len(button_dict) for button_dict in funcs_map.values())

    # Calculate window dimensions based on button count

    # Calculate the width needed for the buttons
    # Each button is width=10 characters plus padding
    button_width = 10 * 8  # 10 characters * approx 8 pixels per character
    padding_width = 10    # 5 pixels padding on each side
    total_button_width = (button_width + padding_width) * min(max_buttons, max_cols)

    # Add some extra width for window borders
    window_width = total_button_width + 40

    # After all tabs and buttons are created, resize the window to fit content
    # Wait a bit for everything to be drawn
    root.update_idletasks()

    # Get the required height based on the content
    required_height = main_frame.winfo_reqheight() + 10  # Add a small padding

    # Set the window size based on content
    root.geometry(f"{window_width}x{required_height}")

    # Start the main loop
    root.mainloop()


if __name__ == "__main__":
    # This is just for debugging - not the entry point
    # Create a mock caller for testing
    wb = xw.Book()
    wb.set_mock_caller()

    # Start the main application
    # Command-line arguments will be automatically processed in the main function
    main()

    # To test with specific tabs enabled, run:
    # python myproject.py --tabs=Market~Portfolio
