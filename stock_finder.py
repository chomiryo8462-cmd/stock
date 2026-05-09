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
        # 1. 데이터 수집
        df = fdr.StockListing('KRX')
        
        # [핵심] 등락률 항목 이름을 자동으로 찾아내는 로직
        possible_cols = ['ChangesRatio', 'ChgRate', '등락률', 'rate']
        target_col = next((c for c in possible_cols if c in df.columns), None)
        
        if target_col is None:
            # 이름을 못 찾으면 가장 비슷한 수치라도 사용합니다.
            target_col = df.columns[5] 

        # 2. 종합 점수 계산 (데이터가 숫자가 아닌 경우를 대비해 변환 추가)
        df[target_col] = pd.to_numeric(df[target_col], errors='coerce').fillna(0)
        df['Volume'] = pd.to_numeric(df['Volume'], errors='coerce').fillna(0)
        
        df['종합점수'] = (df[target_col] * 0.5) + (df['Volume'].rank(pct=True) * 10)
        df['종합점수'] = df['종합점수'].round(2)

        # 3. 파일 생성
        today = datetime.now().strftime('%Y%m%d')
        filename = f"주식분석_{today}.xlsx"
        df.sort_values(by='종합점수', ascending=False).head(100).to_excel(filename, index=False)

        # 4. 메일 발송
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')

        msg = MIMEMultipart()
        msg['Subject'] = f"📊 드디어 완성! 주식 종합 리포트 ({today})"
        msg['To'] = email_user
        msg['From'] = email_user

        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.sendmail(email_user, email_user, msg.as_string())
        
        print("✅ 성공! 보낸 편지함과 받은 편지함을 확인하세요.")

    except Exception as e:
        print(f"❌ 예상치 못한 오류: {str(e)}")

if __name__ == "__main__":
    send_email()
