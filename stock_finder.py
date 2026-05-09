import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

# 1. 주식 데이터 수집 및 분석
df_krx = fdr.StockListing('KRX')
# ... (분석 로직) ...
filename = f"주식리포트_{datetime.now().strftime('%Y%m%d')}.xlsx"
df_krx.head(50).to_excel(filename, index=False)

# 2. 메일 발송
msg = MIMEMultipart()
msg['Subject'] = f"오늘의 주식 리포트 도착 ({datetime.now().strftime('%Y-%m-%d')})"
msg['To'] = os.environ.get('EMAIL_USER')
msg['From'] = os.environ.get('EMAIL_USER')

with open(filename, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={filename}")
    msg.attach(part)

s = smtplib.SMTP('smtp.gmail.com', 587)
s.starttls()
s.login(os.environ.get('EMAIL_USER'), os.environ.get('EMAIL_PASSWORD'))
s.sendmail(msg['From'], msg['To'], msg.as_string())
s.quit()
print("메일 전송 성공!")
