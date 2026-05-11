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
        print("🚀 데이터 수집 시작...")
        # 1. KRX 전체 종목 리스트 확보
        df = fdr.StockListing('KRX')
        
        # 2. 필수 컬럼 확보 및 이름 고정 (에러 방지용)
        # FinanceDataReader의 버전에 따라 컬럼명이 다를 수 있어 강제로 매칭합니다.
        check_cols = {
            'Name': '종목명', 'Close': '현재가', 'ChgRate': '등락률(%)',
            'PER': 'PER', 'PBR': 'PBR', 'DividendYield': '배당수익률(%)',
            'Volume': '거래량', 'Marcap': '시가총액', 'Market': '시장'
        }
        
        final_df = pd.DataFrame()
        for eng, kor in check_cols.items():
            if eng in df.columns:
                final_df[kor] = df[eng]
            else:
                # 컬럼이 없으면 0으로 채운 빈 컬럼 생성
                final_df[kor] = 0

        # 3. 데이터 숫자 변환 (전처리)
        num_cols = ['현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액']
        for col in num_cols:
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0)

        # 4. 주도 테마 분류
        def classify_theme(name):
            if any(k in name for k in ['삼성전자', 'SK하이닉스', '한미반도체', 'HBM']): return 'AI/반도체'
            if any(k in name for k in ['제룡전기', '효성중공업', '현대일렉트릭', '두산에너빌리티']): return '전력/에너지'
            if any(k in name for k in ['대한광통신', '오이솔루션']): return '광통신/5G'
            if any(k in name for k in ['YG', '하이브', '금융', '지주']): return '엔터/금융'
            return '기타'
        
        final_df['주도테마'] = final_df['종목명'].apply(classify_theme)

        # 5. 종합점수 계산 (시총 50% + 거래량 50%)
        final_df['종합점수'] = (final_df['시가총액'].rank(pct=True) * 50 + final_df['거래량'].rank(pct=True) * 50).round(2)
        final_df['시가총액(억)'] = (final_df['시가총액'] / 100000000).astype(int)

        # 6. 상위 100개 추출 및 정리
        cols_to_show = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액(억)', '시장']
        result_df = final_df[cols_to_show].sort_values(by='종합점수', ascending=False).head(100)

        # 7. 엑셀 저장 및 메일 발송
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Master_Report_{today}.xlsx"
        result_df.to_excel(filename, index=False, engine='openpyxl')

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [마스터리포트] 가치지표 및 테마 분석 ({today})"
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
        
        print("✅ 전송 완료!")

    except Exception as e:
        print(f"❌ 오류: {str(e)}")

if __name__ == "__main__":
    send_email()
