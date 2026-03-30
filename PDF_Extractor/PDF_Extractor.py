import os
import smtplib
from email.message import EmailMessage
from datetime import datetime
import shutil

# === CONFIGURATION ===
source_folder = r"C:\Users\sureshp\Desktop\BH-ERP\Python\PDF_Extractor\Archive"  # Network archive folder
extract_folder = r"C:\Users\sureshp\Desktop\BH-ERP\Python\PDF_Extractor\Extract_PDF"  # Local folder to store extracted PDFs
smtp_server = "smtp.bucherhydraulics.com"
smtp_port = 25
from_email = "einvoice.bhre@bucherhydraulics.com"
to_email = "invoice.bhre@bucherhydraulics.com"
subject = "Please See Attached Files"
log_file_path = os.path.join(extract_folder, "email_log.txt")

# === PDF FILE NAMES ===
pdf_names = [
    "IT00657441200_25010731.pdf",
    "IT01700140989_25010730.pdf",
    "IT02591050360_25010729.pdf",
    "IT02591050360_25010728.pdf",
    "IT00657441200_25010727.pdf"

]

# === CREATE EXTRACT FOLDER IF NOT EXISTS ===
os.makedirs(extract_folder, exist_ok=True)

# === OPEN LOG FILE ===
with open(log_file_path, "w", encoding="utf-8") as log_file:

    # === PROCESS EACH PDF ===
    for pdf_name in pdf_names:
        source_path = os.path.join(source_folder, pdf_name)
        destination_path = os.path.join(extract_folder, pdf_name)

        if os.path.exists(source_path):
            # Copy file to extract folder
            shutil.copy2(source_path, destination_path)
            log_file.write(f"{datetime.now()} - Extracted: {pdf_name}\n")

            # Prepare and send email
            try:
                msg = EmailMessage()
                msg["From"] = from_email
                msg["To"] = to_email
                msg["Subject"] = subject
                msg.set_content("Please see attached files.")

                with open(destination_path, "rb") as f:
                    file_data = f.read()
                    msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=pdf_name)

                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.send_message(msg)

                log_file.write(f"{datetime.now()} - PDF: {pdf_name} --> Email sent to: {to_email} - Status: Success\n\n")
                print(f"Email sent: {pdf_name}")
            except Exception as e:
                log_file.write(f"{datetime.now()} - PDF: {pdf_name} --> Email failed to: {to_email} - Error: {str(e)}\n\n")
                print(f"Failed to send {pdf_name}: {e}")
        else:
            log_file.write(f"{datetime.now()} - File not found: {pdf_name}\n")
            print(f"File not found: {pdf_name}")

