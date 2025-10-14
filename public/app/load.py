import win32com.client
import os

def get_outlook_emails():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        emails = [account.SmtpAddress for account in namespace.Accounts]
        return emails if emails else ["Unknown User"]
    except Exception:
        return ["Unknown User"]

if __name__ == "__main__":
    emails = get_outlook_emails()
    
    # Define the path
    assets_dir = r"C:\DSQ Enterprise\assets"
    os.makedirs(assets_dir, exist_ok=True)
    file_path = os.path.join(assets_dir, "USER.txt")
    
    # Overwrite or create the file
    with open(file_path, "w") as f:
        for email in emails:
            f.write(email + "\n")
    
    # Optional: print emails for logging
    for email in emails:
        print(email)
