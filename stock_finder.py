import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

try:
    # 1. 데이터 가져오기
    df = fdr.StockListing('KRX')

    # 2. 종합 점수 및 한글 변환 (이전 분석 로직 복구)
    # 점수 계산을 위해 필요한 수치들 정리 (예시: 거래량과 변동성 기준)
    df['종합점수'] = (df['ChangesRatio'] * 0.5) + (df['Volume'].rank(pct=True) * 10)
    df['종합점수'] = df['종합점수'].round(2)

    # 영어 이름을 한글로 바꾸기
    column_maps = {
        'Code': '종목코드',
        'Name': '종목명',
        'Market': '시장',
        'Close': '현재가',
        'Changes': '대비',
        'ChangesRatio': '등락률',
        'Open': '시가',
        'High': '고가',
        'Low': '저가',
        'Volume': '거래량',
        'Amount': '거래대금',
        'MarCap': '시가총액',
        'Stocks': '상장주식수'
    }
    df = df.rename(columns=column_maps)

    # 점수 높은 순으로 정렬 후 상위 100개만 남기기
    df = df.sort_values(by='종합점수', ascending=False)
    
    # 3. 파일 만들기
    today = datetime.now().strftime('%Y%m%d')
    filename = f"주식_종합분석_{today}.xlsx"
    df.head(100).to_excel(filename, index=False, engine='openpyxl')

    # 4. 이메일 보내기
    email_user = os.environ.get('EMAIL_USER')
    email_password = os.environ.get('EMAIL_PASSWORD')

    msg = MIMEMultipart()
    msg['Subject'] = f"🚀 종합 점수 포함! 주식 리포트 ({today})"
    msg['To'] = email_user
    msg['From'] = email_user

    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={filename}")
        msg.attach(part)

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(email_user, email_password)
        server.sendmail(email_user, email_user, msg.as_string())
    
    print("한글 리포트 발송 완료!")

except Exception as e:
    print(f"오류 발생: {e}")
