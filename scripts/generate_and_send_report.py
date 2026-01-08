import pandas as pd
import smtplib
from email.message import EmailMessage
import os

# ==================================================
# FILE PATHS
# ==================================================
SBI_FILE = "/Users/apple/Downloads/SBI Collection account bank statement_1221.xlsx"
BOOK_FILE = "/Users/apple/Downloads/Book entry pertains to SBI Collection account_1221.xlsx"
OUTPUT_FILE = "/Users/apple/Downloads/SBI_BOOK_Concatenated_Output.xlsx"

# ==================================================
# EMAIL CONFIG
# ==================================================
SENDER_EMAIL = "vatsavs906@gmail.com"
RECEIVER_EMAIL = "tejmaddali22@gmail.com"
APP_PASSWORD = "vujg hmbc xrvz vesv"   # Gmail App Password

# ==================================================
# STEP 1: READ EXCEL FILES
# ==================================================
sbi = pd.read_excel(SBI_FILE, engine="openpyxl")
book = pd.read_excel(BOOK_FILE, engine="openpyxl")

# ==================================================
# STEP 2: ADD SOURCE COLUMN
# ==================================================
sbi["source"] = "SBI"
book["source"] = "BOOK"

# ==================================================
# STEP 3: CONCATENATE DATAFRAMES
# ==================================================
final_df = pd.concat([sbi, book], axis=0, ignore_index=True)

print("Final Columns:", final_df.columns.tolist())
print("Total Columns:", len(final_df.columns))
print("Total Rows:", len(final_df))

# ==================================================
# STEP 4: SAVE OUTPUT FILE
# ==================================================
final_df.to_excel(OUTPUT_FILE, index=False)
print(f"\n✅ Output file generated at:\n{OUTPUT_FILE}")

# ==================================================
# STEP 5: CREATE EMAIL
# ==================================================
msg = EmailMessage()
msg["Subject"] = "SBI & Book Concatenated Report"
msg["From"] = SENDER_EMAIL
msg["To"] = RECEIVER_EMAIL

msg.set_content(
    "Hi,\n\nPlease find attached the concatenated SBI & Book report.\n\nRegards,\nSrivatsav"
)

# ==================================================
# STEP 6: ATTACH FILE
# ==================================================
with open(OUTPUT_FILE, "rb") as f:
    file_data = f.read()
    file_name = os.path.basename(OUTPUT_FILE)

msg.add_attachment(
    file_data,
    maintype="application",
    subtype="octet-stream",
    filename=file_name
)

# ==================================================
# STEP 7: SEND EMAIL
# ==================================================
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(SENDER_EMAIL, APP_PASSWORD)
    server.send_message(msg)

print("✅ Email sent successfully with attachment!")

