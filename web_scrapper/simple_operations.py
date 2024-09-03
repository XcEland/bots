from RPA.Excel.Files import Files
from robocorp.tasks import task

excel = Files()

@task
def minimal_task():
    account_calculation()

def account_calculation():
    try:
        # Create a new workbook and add a worksheet
        excel.create_workbook("accounting.xlsx")
        excel.append_rows_to_worksheet([["Item", "Price", "Quantity", "Total"]], header=False)

        # Data to be added to the worksheet
        data = [
            ["Apple", 1.2, 10],
            ["Banana", 0.8, 15],
            ["Orange", 1.5, 12]
        ]

        excel.append_rows_to_worksheet(data)

        # Apply formulas to calculate totals
        for i in range(2, len(data) + 2):
            formula = f"=B{i}*C{i}"
            excel.set_cell_formula(f"D{i}", formula)

        # Apply formula to calculate the sum of totals
        sum_formula = f"=SUM(D2:D{len(data) + 1})"
        excel.set_cell_formula(f"D{len(data) + 2}", sum_formula)

        # Save the workbook with results
        excel.save_workbook("sample_accounting.xlsx")
        
    except FileNotFoundError:
        print("The specified file could not be found.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure the workbook is closed even if an error occurs
        excel.close_workbook()

