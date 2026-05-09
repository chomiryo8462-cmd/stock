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
        
        # 등락률 항목 자동 찾기
        possible_cols = ['ChangesRatio', 'ChgRate', '등락률', 'rate']
        target_col = next((c for c in possible_cols if c in df.columns), None)
        
        if target_col:
            df[target_col] = pd.to_numeric(df[target_col], errors='coerce').fillna(0)
            df['종합점수'] = (df[target_col] * 0.5) + (df['Volume'].rank(pct=True) * 10)
            df['종합점수'] = df['종합점수'].round(2)
            
            # 한글 이름으로 깔끔하게 정리
            column_maps = {
                'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
                'Close': '현재가', 'Changes': '대비', target_col: '등락률',
                'Volume': '거래량', 'Amount': '거래대금', 'MarCap': '시가총액'
            }
            df = df.rename(columns=column_maps)
            df = df.sort_values(by='종합점수', ascending=False)

        # 2. 파일 생성 (이름에 .xlsx를 확실히 붙이고 엔진 고정)
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Report_{today}.xlsx"
        
        # index=False로 설정하여 불필요한 번호 열을 제거합니다.
        df.head(100).to_excel(filename, index=False, engine='openpyxl')

        # 3. 메일 설정
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')

        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [확인요망] 오늘의 주식 리포트 ({today})"
        msg['To'] = email_user
        msg['From'] = email_user

        # 4. 파일 첨부 (파일 형식을 엑셀로 정확히 지정)
        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        # 5. 전송
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.sendmail(email_user, email_user, msg.as_string())
        
        print(f"✅ {filename} 발송 성공!")

    except Exception as e:
        print(f"❌ 오류: {str(e)}")

if __name__ == "__main__":
    send_email()
