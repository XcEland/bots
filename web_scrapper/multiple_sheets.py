from RPA.Excel.Files import Files

excel = Files()

def multi_sheets_calculations():
    try:
        # Open an existing workbook and set the active worksheet
        excel.open_workbook("orders.xlsx")
        
        # Add data to Sheet1
        excel.set_active_worksheet("Sheet1")
        excel.append_rows_to_worksheet([["Category", "Amount", "Date"]],
                                       header=False)
        expenses = [
            ["Rent", 1200, "2024-09-01"],
            ["Groceries", 300, "2024-09-02"],
            ["Utilities", 150, "2024-09-03"],
            ["Transportation", 100, "2024-09-04"]
        ]
        excel.append_rows_to_worksheet(expenses)
        
        # Add headers to Sheet2 and copy data from Sheet1
        excel.append_rows_to_worksheet([["Category", "Adjusted Amount"]],
                                       header=False, name="Sheet2")

        for i in range(2, len(expenses) + 2):
            # Copy Category from Sheet1 to Sheet2
            category = excel.get_cell_value(i, "A", name="Sheet1")
            excel.set_cell_value(i, "A", category, name="Sheet2")
            
            # Get original amount from Sheet1 and calculate adjusted amount
            original_amount = excel.get_cell_value(i, "B", name="Sheet1")
            adjusted_amount = f"={original_amount}+50"
            excel.set_cell_value(i, "B", adjusted_amount, name="Sheet2")
        
        # Save the workbook with results
        excel.save_workbook("advanced_accounting.xlsx")
        
    except FileNotFoundError:
        print("The specified file could not be found.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure the workbook is closed even if an error occurs
        excel.close_workbook()
