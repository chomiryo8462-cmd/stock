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
        print("🚀 고정밀 데이터 수집 시작...")
        
        # 1. 종목 리스트 확보 (최대한 많은 정보를 가져오기 위해 'KRX' 사용)
        df_krx = fdr.StockListing('KRX')
        
        # 2. 데이터가 0인 문제를 해결하기 위해 컬럼명 수동 매칭 강화
        # 최근 fdr은 'PER', 'PBR', 'DividendYield' 등을 제공하므로 대소문자 무시하고 찾습니다.
        mapping = {
            'Name': '종목명', 'Close': '현재가', 'ChgRate': '등락률(%)',
            'PER': 'PER', 'PBR': 'PBR', 'DividendYield': '배당수익률(%)',
            'Volume': '거래량', 'Marcap': '시가총액', 'Market': '시장'
        }
        
        # 실제 존재하는 컬럼들만 골라서 한글로 변경
        existing_cols = {c: mapping[c] for c in mapping.keys() if c in df_krx.columns}
        df = df_krx[list(existing_cols.keys())].rename(columns=existing_cols)

        # 3. 필수 지표가 없을 경우를 대비한 수치 보정
        for col in ['PER', 'PBR', '배당수익률(%)', '등락률(%)']:
            if col not in df.columns:
                df[col] = 0.0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        # 4. 주도 테마 분류 (사용자 요청 반영)
        def classify_theme(name):
            if any(k in name for k in ['삼성전자', 'SK하이닉스', '한미반도체']): return 'AI/반도체'
            if any(k in name for k in ['제룡전기', '효성중공업', '현대일렉트릭', '두산에너빌리티']): return '전력/에너지'
            if any(k in name for k in ['대한광통신', '오이솔루션']): return '광통신/5G'
            if any(k in name for k in ['YG', '하이브', '금융', '신한']): return '엔터/금융'
            return '기타'
        
        df['주도테마'] = df['종목명'].apply(classify_theme)

        # 5. 종합점수 및 시총 계산
        df['시가총액'] = pd.to_numeric(df['시가총액'], errors='coerce').fillna(0)
        df['거래량'] = pd.to_numeric(df['거래량'], errors='coerce').fillna(0)
        df['종합점수'] = (df['시가총액'].rank(pct=True) * 50 + df['거래량'].rank(pct=True) * 50).round(2)
        df['시가총액(억)'] = (df['시가총액'] / 100000000).astype(int)

        # 6. 최종 엑셀 데이터 구성 (상위 100개)
        target_cols = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액(억)', '시장']
        df_final = df[target_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 7. 엑셀 파일 생성
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Master_Report_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='openpyxl')

        # 8. 메일 전송
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [최종보정] 마스터리포트 ({today})"
        msg['To'] = email_user
        msg['From'] = email_user

        with open(filename, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.send_object = server.sendmail(email_user, email_user, msg.as_string())
        
        print("✅ 리포트 전송 성공!")

    except Exception as e:
        print(f"❌ 오류 발생: {e}")

if __name__ == "__main__":
    send_email()
