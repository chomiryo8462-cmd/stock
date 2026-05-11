import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import subprocess
import sys

# 필수 라이브러리 강제 설치 (에러 방지)
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

try:
    from pykrx import stock
except:
    install('pykrx')
    from pykrx import stock

def send_email():
    try:
        print("🚀 고정밀 데이터(pykrx) 수집 시작...")
        today_str = datetime.now().strftime('%Y%m%d')
        
        # 1. 투자지표(PER/PBR/배당수익률) 가져오기
        # 최근 영업일 기준으로 데이터를 가져옵니다.
        df_invest = stock.get_market_fundamental_by_ticker(today_str, market="ALL")
        # 2. 가격 및 시가총액/거래량 가져오기
        df_price = stock.get_market_ohlcv_by_ticker(today_str, market="ALL")
        
        # 3. 데이터 합치기
        df = pd.concat([df_price, df_invest], axis=1)
        
        # 종목명 추가
        names = [stock.get_market_ticker_name(ticker) for ticker in df.index]
        df['종목명'] = names
        
        # 4. 필요한 컬럼 정리 및 한글화
        df = df.reset_index()
        df = df.rename(columns={
            '종가': '현재가', '등락률': '등락률(%)', '거래량': '거래량',
            'PER': 'PER', 'PBR': 'PBR', '주당배당금': '배당금', '시가총액': '시가총액'
        })

        # 5. 테마 분류
        def classify_theme(name):
            if any(k in name for k in ['삼성전자', 'SK하이닉스', '한미반도체']): return 'AI/반도체'
            if any(k in name for k in ['제룡전기', '효성중공업', '현대일렉트릭', '두산에너빌리티']): return '전력/에너지'
            if any(k in name for k in ['대한광통신', '오이솔루션']): return '광통신/5G'
            if any(k in name for k in ['YG', '하이브', '금융']): return '엔터/금융'
            return '기타'
        
        df['주도테마'] = df['종목명'].apply(classify_theme)

        # 6. 종합점수 계산 및 정리
        df['종합점수'] = (df['시가총액'].rank(pct=True) * 50 + df['거래량'].rank(pct=True) * 50).round(2)
        df['시가총액(억)'] = (df['시가총액'] / 100000000).astype(int)
        
        # 배당수익률 계산 (배당금 / 현재가 * 100)
        df['배당수익률(%)'] = (df['배당금'] / df['현재가'] * 100).round(2)

        cols = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액(억)']
        df_final = df[cols].sort_values(by='종합점수', ascending=False).head(100)

        # 7. 엑셀 저장
        filename = f"Stock_Master_Report_{today_str}.xlsx"
        df_final.to_excel(filename, index=False)

        # 8. 메일 발송
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [완전수정] 마스터리포트 ({today_str})"
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
        
        print("✅ 전송 성공!")
    except Exception as e:
        print(f"❌ 에러: {e}")

if __name__ == "__main__":
    send_email()
