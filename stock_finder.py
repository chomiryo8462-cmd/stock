import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

def send_email():
    try:
        # 1. 데이터 수집 (안전한 방식으로 변경)
        df = fdr.StockListing('KRX')
        
        # 'ChangesRatio'가 없을 경우를 대비해 'ChgPct' 등 다른 이름 확인 및 처리
        # 최신 버전에서는 'ChgRate' 혹은 'ChangesRatio'를 사용합니다.
        target_col = 'ChangesRatio' if 'ChangesRatio' in df.columns else 'ChgRate'
        
        # 2. 종합 점수 계산 (오류 방지 로직 추가)
        df['종합점수'] = (df[target_col] * 0.5) + (df['Volume'].rank(pct=True) * 10)
        df['종합점수'] = df['종합점수'].round(2)

        # 한글 이름으로 변환
        column_maps = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Close': '현재가', 'Changes': '대비', target_col: '등락률',
            'Volume': '거래량', 'Amount': '거래대금', 'MarCap': '시가총액'
        }
        df = df.rename(columns=column_maps)
        df = df.sort_values(by='종합점수', ascending=False)

        # 3. 파일 생성
        today = datetime.now().strftime('%Y%m%d')
        filename = f"주식분석_{today}.xlsx"
        df.head(100).to_excel(filename, index=False, engine='openpyxl')

        # 4. 메일 설정 및 발송 (보안 강화)
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')

        msg = MIMEMultipart()
        msg['Subject'] = f"✅ 오류 해결! 오늘의 주식 리포트 ({today})"
        msg['To'] = email_user
        msg['From'] = email_user

        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        # SSL 방식을 사용하여 더 안전하게 전송
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.sendmail(email_user, email_user, msg.as_string())
        
        print("✅ 드디어 성공! 메일을 확인하세요.")

    except Exception as e:
        print(f"❌ 또 다른 오류 발생: {str(e)}")

if __name__ == "__main__":
    send_email()
