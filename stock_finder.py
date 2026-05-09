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
        # 1. 데이터 가져오기
        df = fdr.StockListing('KRX')
        
        # 2. 등락률 항목 찾기 (영어/한글 모두 대응)
        possible_cols = ['ChangesRatio', 'ChgRate', '등락률', 'rate']
        target_col = next((c for c in possible_cols if c in df.columns), None)

        # 3. 점수 계산 로직 (데이터가 있을 때만 실행)
        if target_col:
            df[target_col] = pd.to_numeric(df[target_col], errors='coerce').fillna(0)
            df['Volume'] = pd.to_numeric(df['Volume'], errors='coerce').fillna(0)
            
            # 종합점수 계산 후 열 추가
            df['종합점수'] = (df[target_col] * 0.5) + (df['Volume'].rank(pct=True) * 10)
            df['종합점수'] = df['종합점수'].round(2)
        else:
            df['종합점수'] = 0

        # 4. 한글 이름표 강제 부여
        # 원본 컬럼 이름이 무엇이든 강제로 우리가 원하는 한글로 바꿉니다.
        rename_dict = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Close': '현재가', 'Changes': '대비', 'Volume': '거래량', 
            'Amount': '거래대금', 'MarCap': '시가총액'
        }
        # 등락률은 찾은 컬럼명을 사용
        if target_col:
            rename_dict[target_col] = '등락률'
            
        df = df.rename(columns=rename_dict)

        # 5. 필요한 항목만 추출 및 정렬 (종합점수 우선)
        display_cols = ['종합점수', '종목명', '현재가', '등락률', '거래량', '시가총액']
        # 실제 존재하는 컬럼만 골라내기
        final_cols = [c for c in display_cols if c in df.columns]
        df_final = df[final_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 6. 엑셀 파일 저장
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Report_{today}.xlsx"
        # 모바일 호환성이 좋은 openpyxl 엔진 사용
        df_final.to_excel(filename, index=False, engine='openpyxl')

        # 7. 메일 보내기
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')

        msg = MIMEMultipart()
        msg['Subject'] = f"🚀 [최종완성] 한글 주식 리포트 ({today})"
        msg['To'] = email_user
        msg['From'] = email_user

        with open(filename, "rb") as attachment:
            # 엑셀 파일임을 명시하는 표준 타입 사용
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.sendmail(email_user, email_user, msg.as_string())
        
        print("✅ 모든 미션 완료! 메일을 확인하세요.")

    except Exception as e:
        print(f"❌ 최종 오류 발생: {str(e)}")

if __name__ == "__main__":
    send_email()
