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
        # 1. 기본 종목 정보와 상세 지표(PER/PBR/배당 등) 수집
        # 오늘 날짜 기준으로 상장 종목의 상세 가치지표 데이터를 가져옵니다.
        df_krx = fdr.StockListing('KRX')
        
        # 2. 숫자 데이터 정제
        numeric_cols = ['Close', 'ChgRate', 'Volume', 'MarCap']
        for col in numeric_cols:
            if col in df_krx.columns:
                df_krx[col] = pd.to_numeric(df_krx[col], errors='coerce').fillna(0)

        # 3. 대세 테마 분류 (AI/반도체, 전력, 광통신 등)
        def detect_theme(name):
            themes = {
                'AI/반도체': ['삼성전자', 'SK하이닉스', '한미반도체', 'HBM', '엔비디아', 'AI'],
                '전력/에너지': ['제룡전기', '효성중공업', 'HD현대일렉트릭', '두산에너빌리티', '변압기'],
                '광통신/5G': ['대한광통신', '오이솔루션', '광통신', '5G'],
                '엔터/문화': ['YG', '하이브', 'JYP', '에스엠']
            }
            for theme, keywords in themes.items():
                if any(key in name for key in keywords): return theme
            return '일반'
        
        df_krx['주도테마'] = df_krx['Name'].apply(detect_theme)

        # 4. 종합점수 계산 (시총 50 + 거래량 50)
        df_krx['종합점수'] = (df_krx['MarCap'].rank(pct=True) * 50) + (df_krx['Volume'].rank(pct=True) * 50)
        df_krx['종합점수'] = df_krx['종합점수'].round(2)

        # 5. 항목 한글화 및 단위 조정
        df_krx['시가총액(억)'] = (df_krx['MarCap'] / 100000000).astype(int)
        
        rename_dict = {
            'Name': '종목명', 'Close': '현재가', 'ChgRate': '등락률(%)',
            'Volume': '거래량', 'Market': '시장'
        }
        # PER, PBR 등이 데이터에 있으면 한글로 변경
        for col in ['PER', 'PBR', 'DividendYield']:
            if col in df_krx.columns:
                rename_dict[col] = '배당수익률' if col == 'DividendYield' else col

        df_krx = df_krx.rename(columns=rename_dict)

        # 6. 리포트 구성 (상위 100개)
        # 존재하지 않는 항목이 있어도 에러 없이 진행되도록 필터링
        desired_cols = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률', '거래량', '시가총액(억)', '시장']
        final_cols = [c for c in desired_cols if c in df_krx.columns]
        
        df_final = df_krx[final_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 7. 엑셀 저장 및 메일 발송
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Master_Report_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='openpyxl')

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        msg = MIMEMultipart()
        msg['Subject'] = f"🏆 [마스터리포트] 주도테마 및 가치지표 분석 ({today})"
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
        
        print("✅ 테마 및 가치지표 포함 리포트 전송 성공!")

    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    send_email()
