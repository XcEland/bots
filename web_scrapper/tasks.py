from robocorp.tasks import task
from RPA.Excel.Files import Files
from RPA.Email.ImapSmtp import ImapSmtp
from RPA.Email.Exchange import Exchange
from RPA.Robocorp.Vault import Vault

from robocorp import browser

from datetime import date

excel = Files()


_secret = Vault().get_secret("credentials")
gmail_account = _secret["email"]
gmail_password = _secret["password"]


# vault_name = "email_oauth_microsoft"
# secrets = Vault().get_secret(vault_name)
# ex_account = "mkoma.xyz@gmail.com"

# mail = Exchange(
#     vault_name=vault_name,
#     vault_token_key="token",
#     tenant="ztzvn.onmicrosoft.com"
# )
# mail.authorize(
#     username=ex_account,
#     is_oauth=True,
#     client_id=secrets["client_id"],
#     client_secret=secrets["client_secret"],
#     token=secrets["token"]
# )
# mail.send_message(
#     recipients="findyandx@gmail.com",
#     subject="Message from RPA Python",
#     body="RPA Python message body",
# )

@task
def minimal_task():
    # read_excel_worksheet()
    # send_email()
    # send_batch_emails()
    # send_batch_emails()
    # set_formula()
    # account_calculation()
    # intermediate_accounting()
    # multi_sheets_calculations()
    web_scraper_top_10_crypto()

def read_excel_worksheet():
    workbook = Files()
    workbook.open_workbook("orders.xlsx")
    try:
        worksheet= workbook.read_worksheet_as_table("Sheet1", header=False)
        workbook.set_cell_value(1, 1, "Some value")
        workbook.set_cell_value(1, "C", "Some Other value")
        workbook.save_workbook()
        for row in worksheet:
            print(row)
    finally:
        workbook.close_workbook()

def send_email():
    
    attachment_path_1 = "C:/Users/Admin/Desktop/bots/web_scrapper/complex_accounting.xlsx"
    attachment_path_2 = "C:/Users/Admin/Desktop/bots/web_scrapper/orders.xlsx"

    mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
    mail.authorize(account=gmail_account, password=gmail_password)
    mail.send_message(
    sender=gmail_account,
    recipients="findyandx@gmail.com",
    subject="Message from RPA Python",
    body="RPA Python message body",
    attachments= [attachment_path_1, attachment_path_2],)

def send_emails_from_excel(employee):
    attachment_path_1 = "C:/Users/Admin/Desktop/bots/web_scrapper/complex_accounting.xlsx"
    attachment_path_2 = "C:/Users/Admin/Desktop/bots/web_scrapper/orders.xlsx"

    mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
    mail.authorize(account=gmail_account, password=gmail_password)
    mail.send_message(
    sender=gmail_account,
    recipients=str(employee["email"]),
    subject=str(employee["subject"]),
    body=str(employee["body"]),
    attachments= [attachment_path_1, attachment_path_2],)

def send_batch_emails():
    excel = Files()
    excel.open_workbook("employeeData.xlsx")
    worksheet = excel.read_worksheet_as_table("Sheet1", header=True)
    excel.close_workbook()

    for row in worksheet:
        send_emails_from_excel(row)

def set_formula():
    workbook = Files()
    workbook.open_workbook("orders.xlsx")
    try:
        worksheet= workbook.read_worksheet_as_table("Sheet1", header=True)
        for i in range(2, 6):  # This loop will iterate from 2 to 7
            formula = f"=B{i}+C{i}"
            cell_range = f"E{i}"
            workbook.set_cell_formula(cell_range, formula, True)

        workbook.save_workbook()
        for row in worksheet:
            print(row)
    finally:
        workbook.close_workbook()

def account_calculation():
    excel.create_workbook("accounting.xlsx")
    excel.append_rows_to_worksheet([["Item", "Price", "Quantity", "Total"]],header=False)
    
    data = [
    ["Apple", 1.2, 10],
    ["Banana", 0.8, 15],
    ["Orange", 1.5, 12]
    ]

    excel.append_rows_to_worksheet(data)

    for i in range(2, len(data)+2):
        formula = f"=B{i}*C{i}"
        excel.set_cell_formula(f"D{i}", formula)

    sum_formula = f"=SUM(D2:D{len(data) + 1})"
    excel.set_cell_formula(f"D{len(data) + 2}", sum_formula)

    excel.save_workbook("sample_accounting.xlsx")
    excel.close_workbook()

def intermediate_accounting():
    excel.create_workbook("complex_accounting.xlsx")
    excel.append_rows_to_worksheet([["Item", "Price", "Quantity", "Subtotal", "Tax (10%)", "Total"]],
                               header=False)
    
    data = [
    ["Laptop", 1200, 2],
    ["Mouse", 25, 5],
    ["Keyboard", 45, 3],
    ["Monitor", 300, 1]]

    excel.append_rows_to_worksheet(data)

    # Define formulas for Subtotal, Tax, and Total columns
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

    excel.save_workbook("complex_accounting.xlsx")
    excel.close_workbook()

def multi_sheets_calculations():
    # Create a new Excel file
    excel.open_workbook("orders.xlsx")
   
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

    # Sheet2: Copy the category and amount to another sheet
    excel.append_rows_to_worksheet([["Category", "Adjusted Amount"]],
                               header=False, name="Sheet2")

    for i in range(2, len(expenses) + 2):
        # Copy Category
        category = excel.get_cell_value(i,"A",name="Sheet1")

        # excel.set_active_worksheet("Sheet2")
        excel.set_cell_value(i, "A", category,name="Sheet2")
        
        # Get original amount from Sheet1
        original_amount = excel.get_cell_value(i,"B",name="Sheet1")
        
        # Adjusted Amount = Original Amount + $50
        adjusted_amount = f"={original_amount}+50"
        excel.set_cell_value(i,"B", adjusted_amount,name="Sheet2")
        
        # Save and close the workbook
        excel.save_workbook("advanced_accounting.xlsx")
        excel.close_workbook()

def web_scraper_top_10_crypto() -> None:
    try:
        # initialize the browser and a new webpage
        page = browser.page()
        # moving to Yahoo Finance
        page.goto("https://finance.yahoo.com/crypto")
        page.wait_for_load_state()

        # make sure we get rid of the cookies permission popup
        try:
            reject_cookies = page.locator("[name=reject]")
            reject_cookies.click()
        except Exception:
            pass
        # ignore any issues if the popup doesn't show up

        # wait for the page to actually show the target table
        crypto_table = page.locator(
            "xpath=//span[contains(.,'Matching Cryptocurrencies')]"
        )
        crypto_table.wait_for(timeout=5000, state="visible")
        # making sure the element is visible
        assert crypto_table.is_visible()

        # scrape the web page and extract the top 10 crypto currencies
        print("#" * 50)
        print("### Top 10 Cryptocurrencies:")
        print("#" * 50)
        csv_content = ["Index,Crypto,Value"]
        for index in range(1, 11):
            # get 1st cell in first row - it contains the name of the crypto ticker
            crypto = page.locator(f".simpTblRow:nth-child({index}) > .Px\(10px\)")
            # make sure the element is visible
            crypto.wait_for(timeout=5000, state="visible")
            # making sure the element is visible
            assert crypto.is_visible()
            # grab the value
            crypto_ticker = crypto.inner_text()

            # get 2rd cell in first row - it contains the name of the crypto ticker
            crypto = page.locator(
                f".simpTblRow:nth-child({index}) > .Va\(m\):nth-child(3)"
            )
            crypto.wait_for(timeout=5000, state="visible")
            assert crypto.is_visible()
            crypto_value = crypto.inner_text()

            # save the content to the csv
            csv_content.append(f'{index},{crypto_ticker},"{crypto_value}"')

            # calculate how to space things to have a table look
            i_spaces = 3 - len(str(index))
            tab_spaces = 25 - len(crypto_ticker)
            print(
                f"### {index}{' '*i_spaces}| {crypto_ticker}{' '*tab_spaces}| $ {crypto_value}"
            )
        print("#" * 50)

        # save the CSV file to output folder
        csv_file = f"output/top-10-cryptos-{date.today()}.csv"
        print(f"### Saving to the CSV file: {csv_file}")
        with open(csv_file, mode="w") as csv:
            csv.writelines([line + "\n" for line in csv_content])
        print("### Done!")

    finally:
        browser.context().close()
        browser.browser().close()
    
if __name__=="main":
    minimal_task()
