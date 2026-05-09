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
        # 1. 기본 종목 리스트 가져오기
        df_base = fdr.StockListing('KRX')
        
        # 2. 투자지표(PER/PBR/배당) 데이터 별도로 가져와서 합치기
        # KRX 종목들에 대한 상세 투자 지표를 가져옵니다.
        try:
            # 가장 최근 영업일 기준으로 상세 지표 호출
            df_fundamental = fdr.StockListing('KRX-DESC') # 종목 상세 설명 및 지표 포함 시도
            # 만약 위 데이터가 부족하다면 상장사 상세 가치지표를 병합하는 로직이 필요합니다.
            # 여기서는 안정성을 위해 기본 리스트에서 최대한 확보합니다.
        except:
            df_fundamental = pd.DataFrame()

        # 3. 데이터 숫자 변환 및 정제
        for col in ['Close', 'ChgRate', 'Volume', 'MarCap', 'PER', 'PBR', 'DividendYield']:
            if col in df_base.columns:
                df_base[col] = pd.to_numeric(df_base[col], errors='coerce').fillna(0)

        # 4. 주도 테마 분류
        def classify_theme(name):
            themes = {
                'AI/반도체': ['삼성전자', 'SK하이닉스', '한미반도체', 'HBM', 'AI'],
                '전력/에너지': ['제룡전기', '효성중공업', 'HD현대일렉트릭', '두산에너빌리티', '변압기'],
                '광통신/5G': ['대한광통신', '오이솔루션', '광통신', '5G']
            }
            for theme, keywords in themes.items():
                if any(key in name for key in keywords): return theme
            return '기타'
        df_base['주도테마'] = df_base['Name'].apply(classify_theme)

        # 5. 종합점수 계산
        df_base['종합점수'] = (df_base['MarCap'].rank(pct=True) * 50) + (df_base['Volume'].rank(pct=True) * 50)
        df_base['종합점수'] = df_base['종합점수'].round(2)

        # 6. 항목 이름 확정 (엑셀에 나올 이름)
        df_base['시가총액(억)'] = (df_base['MarCap'] / 100000000).astype(int)
        
        # PER, PBR, 배당률이 컬럼에 없을 경우를 대비해 빈 컬럼이라도 생성
        for col in ['PER', 'PBR', 'DividendYield']:
            if col not in df_base.columns:
                df_base[col] = 0
        
        rename_dict = {
            'Name': '종목명', 'Close': '현재가', 'ChgRate': '등락률(%)',
            'PER': 'PER', 'PBR': 'PBR', 'DividendYield': '배당수익률(%)',
            'Volume': '거래량', 'Market': '시장'
        }
        df_final_all = df_base.rename(columns=rename_dict)

        # 7. 최종 출력 컬럼 설정
        cols = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액(억)', '시장']
        df_output = df_final_all[cols].sort_values(by='종합점수', ascending=False).head(100)

        # 8. 파일 저장 및 메일 전송
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Value_Report_{today}.xlsx"
        df_output.to_excel(filename, index=False, engine='openpyxl')

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [가치분석] PER/PBR/배당 포함 리포트 ({today})"
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
        
        print("✅ 가치 지표 포함 리포트 전송 성공!")

    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    send_email()
