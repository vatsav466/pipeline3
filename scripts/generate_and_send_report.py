import pandas as pd
import smtplib
from email.message import EmailMessage
import os

# ==================================================
# FILE PATHS (RELATIVE PATHS)
# ==================================================
SBI_FILE = "data/SBI Collection account bank statement_1221.xlsx"
BOOK_FILE = "data/Book entry pertains to SBI Collection account_1221.xlsx"
OUTPUT_FILE = "data/SBI_BOOK_Concatenated_Output.xlsx"

# ==================================================
# EMAIL CONFIG (FROM SECRETS)
# ==================================================
SENDER_EMAIL = os.environ["SENDER_EMAIL"]
RECEIVER_EMAIL = os.environ["RECEIVER_EMAIL"]
APP_PASSWORD = os.environ["EMAIL_APP_PASSWORD"]

# ==================================================
# READ EXCEL FILES
# ==================================================
sbi = pd.read_excel(SBI_FILE, engine="openpyxl")
book = pd.read_excel(BOOK_FILE, engine="openpyxl")

# ==================================================
# ADD SOURCE COLUMN
# ==================================================
sbi["source"] = "SBI"
book["source"] = "BOOK"

# ==================================================
# CONCATENATE
# ==================================================
final_df = pd.concat([sbi, book], ignore_index=True)

# ==================================================
# SAVE OUTPUT
# ==================================================
final_df.to_excel(OUTPUT_FILE, index=False)

# ==================================================
# EMAIL
# ==================================================
msg = EmailMessage()
msg["Subject"] = "SBI & Book Concatenated Report"
msg["From"] = SENDER_EMAIL
msg["To"] = RECEIVER_EMAIL

msg.set_content(
    "Hi,\n\nPlease find attached the concatenated SBI & Book report.\n\nRegards,\nSrivatsav"
)

with open(OUTPUT_FILE, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="octet-stream",
        filename="SBI_BOOK_Concatenated_Output.xlsx"
    )

# ==================================================
# SEND EMAIL (TLS – MORE RELIABLE)
# ==================================================
with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls()
    server.login(SENDER_EMAIL, APP_PASSWORD)
    server.send_message(msg)

print("✅ Email sent successfully")
