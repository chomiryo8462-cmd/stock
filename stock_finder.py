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
        
        # 2. 등락률 항목 자동 찾기 (이름이 영어든 한글이든 대응 가능)
        # 등락률은 보통 소수점(%) 데이터가 들어있는 열입니다.
        possible_cols = ['ChangesRatio', 'ChgRate', '등락률', 'rate']
        target_col = next((c for c in possible_cols if c in df.columns), None)

        if target_col:
            # 숫자로 변환 (모바일에서 깨짐 방지)
            df[target_col] = pd.to_numeric(df[target_col], errors='coerce').fillna(0)
            df['Volume'] = pd.to_numeric(df['Volume'], errors='coerce').fillna(0)
            
            # 종합점수 계산
            df['종합점수'] = (df[target_col] * 0.5) + (df['Volume'].rank(pct=True) * 10)
            df['종합점수'] = df['종합점수'].round(2)
            
            # 한글 이름표 붙이기 (오류 방지를 위해 안전하게 처리)
            column_maps = {
                'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
                'Close': '현재가', 'Changes': '대비', target_col: '등락률',
                'Volume': '거래량', 'Amount': '거래대금', 'MarCap': '시가총액'
            }
            df = df.rename(columns=column_maps)

        # 3. 필요한 항목만 추출 및 정렬 (종합점수 우선)
        # 실제 존재하는 컬럼만 선택하도록 안전하게 필터링
        display_cols = ['종합점수', '종목명', '현재가', '등락률', '거래량', '시가총액']
        final_cols = [c for c in display_cols if c in df.columns]
        
        df_final = df[final_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 4. 엑셀 파일 생성 (.xlsx)
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Analysis_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='xlsxwriter')

        # 5. 메일 발송 설정
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')

        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [최종] 오늘의 주식 종합 분석 리포트 ({today})"
        msg['To'] = email_user
        msg['From'] = email_user

        # 6. 파일 첨부 (모바일 최적화 형식)
        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        # 7. 전송
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.sendmail(email_user, email_user, msg.as_string())
        
        print("✅ 드디어 모든 오류 해결! 메일함을 확인하세요.")

    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    send_email()
