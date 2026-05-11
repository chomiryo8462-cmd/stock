import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from pykrx import stock

def get_recent_business_day():
    # 오늘부터 거꾸로 가며 데이터가 있는 날을 찾음
    target_date = datetime.now()
    for _ in range(7):
        date_str = target_date.strftime('%Y%m%d')
        df = stock.get_market_ohlcv_by_ticker(date_str, market="KOSPI")
        if not df.empty:
            return date_str
        target_date -= timedelta(days=1)
    return datetime.now().strftime('%Y%m%d')

def send_email():
    try:
        print("🚀 고정밀 데이터 수집 시작...")
        target_date = get_recent_business_day()
        print(f"기준 날짜: {target_date}")

        # 1. 시세 및 기본 지표 수집
        df_price = stock.get_market_ohlcv_by_ticker(target_date, market="ALL")
        df_fund = stock.get_market_fundamental_by_ticker(target_date, market="ALL")
        
        # 2. 데이터 병합
        df = pd.concat([df_price, df_fund], axis=1)
        df['종목명'] = [stock.get_market_ticker_name(t) for t in df.index]
        df = df.reset_index().rename(columns={'티커': '코드', '종가': '현재가', '시가총액': '시총'})

        # 3. 테마 분류 로직
        def classify_theme(name):
            if any(k in name for k in ['삼성전자', 'SK하이닉스', '한미반도체']): return 'AI/반도체'
            if any(k in name for k in ['제룡전기', '효성중공업', '현대일렉트릭', '두산에너빌리티']): return '전력/에너지'
            if any(k in name for k in ['대한광통신', '오이솔루션']): return '광통신/5G'
            if any(k in name for k in ['YG', '하이브', '금융']): return '엔터/금융'
            return '기타'
        
        df['주도테마'] = df['종목명'].apply(classify_theme)

        # 4. 종합점수 및 데이터 정리
        df['거래량'] = pd.to_numeric(df['거래량'], errors='coerce').fillna(0)
        df['시총'] = pd.to_numeric(df['시총'], errors='coerce').fillna(0)
        df['종합점수'] = (df['시총'].rank(pct=True) * 50 + df['거래량'].rank(pct=True) * 50).round(2)
        df['시가총액(억)'] = (df['시총'] / 100000000).astype(int)
        
        # PER, PBR 등이 0인 경우 처리 (BPS, EPS 등이 없으면 0)
        for col in ['PER', 'PBR', 'DIV']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                df[col] = 0

        cols = ['종합점수', '주도테마', '종목명', '현재가', '등락률', 'PER', 'PBR', 'DIV', '거래량', '시가총액(억)']
        df_final = df[cols].sort_values(by='종합점수', ascending=False).head(100)
        df_final.columns = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액(억)']

        # 5. 엑셀 저장 및 메일 발송
        filename = f"Stock_Report_{target_date}.xlsx"
        df_final.to_excel(filename, index=False)

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [최종완성] 주식 마스터 리포트 ({target_date})"
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
            server.sendmail(email_user, email_user, msg.as_string())
        
        print(f"✅ {target_date} 데이터 전송 완료!")
    except Exception as e:
        print(f"❌ 에러 발생: {e}")

if __name__ == "__main__":
    send_email()
