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
        
        # 2. 한글로 이름표 미리 바꾸기 (이 작업이 먼저 되어야 합니다)
        column_maps = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Close': '현재가', 'ChangesRatio': '등락률', 'ChgRate': '등락률',
            'Volume': '거래량', 'Amount': '거래대금', 'MarCap': '시가총액'
        }
        df = df.rename(columns=column_maps)

        # 3. 종합점수 계산 및 정렬 (모바일 가독성을 위해 숫자로 변환)
        target_col = '등락률'
        df[target_col] = pd.to_numeric(df[target_col], errors='coerce').fillna(0)
        df['거래량'] = pd.to_numeric(df['거래량'], errors='coerce').fillna(0)
        
        df['종합점수'] = (df[target_col] * 0.5) + (df['거래량'].rank(pct=True) * 10)
        df['종합점수'] = df['종합점수'].round(2)

        # 필요한 항목만 남기고 순서 정리 (종합점수, 종목명, 현재가 순)
        final_cols = ['종합점수', '종목명', '현재가', '등락률', '거래량', '시가총액', '시장']
        df_final = df[final_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 4. 엑셀 파일 생성 (가장 표준적인 형식)
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Report_{today}.xlsx"
        
        # 엔진을 xlsxwriter로 사용하여 모바일 호환성을 극대화합니다.
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='주식리포트')

        # 5. 메일 발송 설정
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')

        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [완료] 한글 주식 리포트 ({today})"
        msg['To'] = email_user
        msg['From'] = email_user

        # 6. 파일 첨부 (파일 형식을 엑셀 전용으로 강력하게 지정)
        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.sendmail(email_user, email_user, msg.as_string())
        
        print("✅ 한글화 및 모바일 대응 리포트 발송 완료!")

    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    send_email()
