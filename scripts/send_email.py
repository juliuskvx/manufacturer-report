"""
send_email.py — sends the monthly manufacturer PPTX via Gmail SMTP
"""
import os
import glob
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

GMAIL_USER = os.environ["GMAIL_USER"]
GMAIL_PASS = os.environ["GMAIL_APP_PASSWORD"]
REPORT_DATE = os.environ.get("REPORT_DATE", "")

# Load data for subject line
with open("output/manufacturer_data.json") as f:
    d = json.load(f)

B = d["boeing"]
A = d["airbus"]

subject = f"EuroAir Intel — Manufacturer Report {d['reportMonth']} | BA ${B['stock'].get('price','—')} | AIR.PA €{A['stock'].get('price','—')}"

body = f"""EuroAir Intel — Monthly Manufacturer Intelligence Report
{d['reportMonth']} | Data as of {d['asOf']}

BOEING
  Deliveries YTD:  {B['deliveries_ytd']}
  Orders YTD:      {B['orders_ytd']}
  Backlog:         {B['backlog']:,} aircraft ({B['backlog_years']} years)
  BA Stock:        ${B['stock'].get('price','N/A')} ({B['stock'].get('change_pct',0):+.2f}% today)

AIRBUS
  Deliveries YTD:  {A['deliveries_ytd']}
  Orders YTD:      {A['orders_ytd']}
  Backlog:         {A['backlog']:,} aircraft ({A['backlog_years']} years)
  AIR.PA Stock:    €{A['stock'].get('price','N/A')} ({A['stock'].get('change_pct',0):+.2f}% today)

Full analysis in the attached PPTX.

—
EuroAir Intel · aviotiongeek@gmail.com
Published monthly on the 15th
"""

# Find PPTX
files = glob.glob("output/Manufacturer_Intelligence_Report_*.pptx")
if not files:
    print("ERROR: no PPTX found to attach")
    exit(1)
pptx_path = sorted(files)[-1]

msg = MIMEMultipart()
msg["From"]    = GMAIL_USER
msg["To"]      = GMAIL_USER
msg["Subject"] = subject
msg.attach(MIMEText(body, "plain"))

with open(pptx_path, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(pptx_path)}"')
msg.attach(part)

print(f"Sending to {GMAIL_USER}...")
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(GMAIL_USER, GMAIL_PASS)
    server.sendmail(GMAIL_USER, GMAIL_USER, msg.as_string())

print(f"✓ Email sent: {subject}")
