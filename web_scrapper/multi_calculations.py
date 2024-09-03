from RPA.Excel.Files import Files

excel = Files()
def intermediate_accounting():
    try:
        # Create a new workbook and add a worksheet with headers
        excel.create_workbook("complex_accounting.xlsx")
        excel.append_rows_to_worksheet([["Item", "Price", "Quantity", "Subtotal", "Tax (10%)", "Total"]],
                                       header=False)

        # Data to be added to the worksheet
        data = [
            ["Laptop", 1200, 2],
            ["Mouse", 25, 5],
            ["Keyboard", 45, 3],
            ["Monitor", 300, 1]
        ]

        excel.append_rows_to_worksheet(data)

        # Apply formulas for Subtotal, Tax, and Total columns
        for i in range(2, len(data) + 2):
            # Subtotal = Price * Quantity
            subtotal_formula = f"=B{i}*C{i}"
            excel.set_cell_formula(f"D{i}", subtotal_formula)

            # Tax = Subtotal * 10%
            tax_formula = f"=D{i}*0.10"
            excel.set_cell_formula(f"E{i}", tax_formula)

            # Total = Subtotal + Tax
            total_formula = f"=D{i}+E{i}"
            excel.set_cell_formula(f"F{i}", total_formula)

        # Save the workbook with results
        excel.save_workbook("complex_accounting.xlsx")

    except FileNotFoundError:
        print("The specified file could not be found.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure the workbook is closed even if an error occurs
        excel.close_workbook()
