import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

try:
    # 1. 주식 데이터 가져오기 (가장 안정적인 방식)
    df = fdr.StockListing('KRX')
    
    # 2. 파일 이름 정하기 (오늘 날짜)
    today = datetime.now().strftime('%Y%m%d')
    filename = f"Stock_Report_{today}.xlsx"
    
    # 3. 엑셀 파일 만들기 (엔진을 openpyxl로 고정)
    df.head(100).to_excel(filename, index=False, engine='openpyxl')
    print(f"{filename} 파일 생성 완료!")

    # 4. 이메일 보내기 설정
    email_user = os.environ.get('EMAIL_USER')
    email_password = os.environ.get('EMAIL_PASSWORD')

    msg = MIMEMultipart()
    msg['Subject'] = f"🏆 오늘의 주식 분석 리포트 ({today})"
    msg['To'] = email_user
    msg['From'] = email_user

    # 파일 첨부
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={filename}")
        msg.attach(part)

    # 메일 전송
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(email_user, email_password)
        server.sendmail(email_user, email_user, msg.as_string())
    
    print("메일 발송 성공!")

except Exception as e:
    print(f"오류 발생: {e}")
