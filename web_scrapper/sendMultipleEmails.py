from robocorp.tasks import task
from RPA.Excel.Files import Files
from RPA.Email.ImapSmtp import ImapSmtp
from RPA.Robocorp.Vault import Vault

_secret = Vault().get_secret("credentials")
gmail_account = _secret["email"]
gmail_password = _secret["password"]

excel = Files()

# @task
# def minimal_task():
#     send_batch_emails()

def send_emails_from_excel(employee):
    
    try:
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

    except FileNotFoundError:
        print("The Excel file was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")


def send_batch_emails():
    excel = Files()
    try:
        excel.open_workbook("employeeData.xlsx")
        worksheet = excel.read_worksheet_as_table("Sheet1", header=True)
    except FileNotFoundError:
        print("The Excel file was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        excel.close_workbook()

    for row in worksheet:
        send_emails_from_excel(row)
