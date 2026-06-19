"""
send_email.py — sends the monthly manufacturer PPTX + Excel via Gmail SMTP
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

with open("output/manufacturer_data.json") as f:
    d = json.load(f)

B = d["boeing"]
A = d["airbus"]
Bs = B.get("stock", {})
As = A.get("stock", {})

b_chg = float(Bs.get("change_pct", 0))
a_chg = float(As.get("change_pct", 0))

subject = (
    f"EuroAir Intel — Manufacturer Report {d['reportMonth']}  |  "
    f"BA ${Bs.get('price','—')} ({b_chg:+.2f}%)  |  "
    f"AIR.PA €{As.get('price','—')} ({a_chg:+.2f}%)"
)

body = f"""EuroAir Intel — Monthly Manufacturer Intelligence Report
{d['reportMonth']}  ·  Data as of {d['asOf']}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BOEING
  Deliveries YTD:   {B['deliveries_ytd']}
  Gross Orders YTD: {B['orders_ytd']}
  Total Backlog:    {B['backlog']:,} aircraft ({B['backlog_years']} years)
  737 MAX:          {B['models'].get('737 MAX', 0)} delivered YTD
  787 Dreamliner:   {B['models'].get('787', 0)} delivered YTD
  BA Stock:         ${Bs.get('price','N/A')}  ({b_chg:+.2f}% today)
  52w Range:        ${Bs.get('week52_low','N/A')} – ${Bs.get('week52_high','N/A')}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AIRBUS
  Deliveries YTD:   {A['deliveries_ytd']}
  Gross Orders YTD: {A['orders_ytd']}
  Total Backlog:    {A['backlog']:,} aircraft ({A['backlog_years']} years)
  A320neo family:   {A['models'].get('A320neo', 0)} delivered YTD
  A350:             {A['models'].get('A350', 0)} delivered YTD
  AIR.PA Stock:     €{As.get('price','N/A')}  ({a_chg:+.2f}% today)
  52w Range:        €{As.get('week52_low','N/A')} – €{As.get('week52_high','N/A')}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Attachments:
  • Manufacturer_Intelligence_Report_{d['reportDate']}.pptx  (9-slide deck)
  • Manufacturer_Data_{d['reportDate']}.xlsx  (5-sheet data archive)

—
EuroAir Intel · aviotiongeek@gmail.com
Published monthly on the 15th
"""

# Find files
pptx_files = sorted(glob.glob("output/Manufacturer_Intelligence_Report_*.pptx"))
xlsx_files = sorted(glob.glob("output/Manufacturer_Data_*.xlsx"))

attachments = []
if pptx_files:
    attachments.append(pptx_files[-1])
else:
    print("WARNING: no PPTX found")

if xlsx_files:
    attachments.append(xlsx_files[-1])
else:
    print("WARNING: no Excel found")

if not attachments:
    print("ERROR: no files to attach")
    exit(1)

msg = MIMEMultipart()
msg["From"]    = GMAIL_USER
msg["To"]      = GMAIL_USER
msg["Subject"] = subject
msg.attach(MIMEText(body, "plain"))

for path in attachments:
    with open(path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(path)}"')
    msg.attach(part)
    print(f"  Attaching: {os.path.basename(path)}")

print(f"Sending to {GMAIL_USER}...")
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(GMAIL_USER, GMAIL_PASS)
    server.sendmail(GMAIL_USER, GMAIL_USER, msg.as_string())

print(f"✓ Email sent with {len(attachments)} attachments")
print(f"  Subject: {subject}")
