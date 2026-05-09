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
        # 1. 전 종목 데이터 가져오기
        df = fdr.StockListing('KRX')
        
        # 2. 데이터 전처리 (숫자 변환)
        cols_to_fix = ['Close', 'Changes', 'ChgRate', 'Volume', 'Amount', 'MarCap']
        for col in cols_to_fix:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 3. 종합 점수 로직 고도화 (시가총액과 거래량 비중 조절)
        # 등락률이 0인 주말에도 시가총액이 크고 거래가 활발한 종목을 상위에 배치
        df['종합점수'] = (df['MarCap'].rank(pct=True) * 40) + \
                        (df['Volume'].rank(pct=True) * 30) + \
                        (df['ChgRate'] * 30)
        df['종합점수'] = df['종합점수'].round(2)

        # 4. 한글 이름표 및 항목 확장
        rename_dict = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Close': '현재가(원)', 'Changes': '전일대비', 'ChgRate': '등락률(%)',
            'Volume': '거래량', 'Amount': '거래대금', 'MarCap': '시가총액(억)',
            'Stocks': '상장주식수'
        }
        df = df.rename(columns=rename_dict)

        # 5. 최종 리포트 구성 (상위 100개)
        final_cols = ['종합점수', '종목명', '현재가(원)', '등락률(%)', '전일대비', '거래량', '거래대금', '시가총액(억)', '시장']
        df_report = df[final_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 6. 엑셀 파일 저장 (숫자 서식 적용)
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Premium_Report_{today}.xlsx"
        df_report.to_excel(filename, index=False, engine='openpyxl')

        # 7. 메일 전송
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')

        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [프리미엄] 오늘의 주식 종합 분석 리포트 ({today})"
        msg['To'] = email_user
        msg['From'] = email_user

        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.sendmail(email_user, email_user, msg.as_string())
        
        print("✅ 프리미엄 리포트 발송 성공!")

    except Exception as e:
        print(f"❌ 오류: {str(e)}")

if __name__ == "__main__":
    send_email()
