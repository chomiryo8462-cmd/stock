import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from pykrx import stock

def get_recent_date():
    target = datetime.now()
    for _ in range(10):
        d_str = target.strftime('%Y%m%d')
        if not stock.get_market_ohlcv_by_ticker(d_str, market="KOSPI").empty:
            return d_str
        target -= timedelta(days=1)
    return datetime.now().strftime('%Y%m%d')

def send_email():
    try:
        date = get_recent_date()
        print(f"🚀 {date} 데이터 추출 시작...")

        # 1. 시세 및 기본지표 수집
        df_p = stock.get_market_ohlcv_by_ticker(date, market="ALL")
        df_f = stock.get_market_fundamental_by_ticker(date, market="ALL")
        
        # 2. 데이터 병합
        df = pd.concat([df_p, df_f], axis=1)
        df['종목명'] = [stock.get_market_ticker_name(t) for t in df.index]
        
        # 3. 컬럼 위치 기반 추출 (이름이 바뀌어도 작동하도록)
        # 종가(보통 4번째), 거래량(보통 5번째), 시가총액(보통 7~8번째) 등을 안전하게 처리
        # 직접적으로 필요한 값들만 정리
        res = pd.DataFrame(index=df.index)
        res['종목명'] = df['종목명']
        res['현재가'] = df.iloc[:, 3] # Close 위치
        res['거래량'] = df.iloc[:, 4] # Volume 위치
        res['시가총액'] = df.iloc[:, 6] # Marcap 위치
        
        # PER, PBR, DIV 등은 이름으로 찾되 없으면 0 처리
        for c in ['PER', 'PBR', 'DVD_YLD', 'DIV']:
            res[c] = df[c] if c in df.columns else 0

        # 4. 테마 및 점수 계산
        def get_theme(n):
            if any(k in n for k in ['삼성전자', 'SK하이닉스', '한미반도체']): return 'AI/반도체'
            if any(k in n for k in ['제룡전기', '효성중공업', '현대일렉트릭', '두산에너빌리티']): return '전력/에너지'
            if any(k in n for k in ['대한광통신', '오이솔루션']): return '광통신/5G'
            return '기타'

        res['주도테마'] = res['종목명'].apply(get_theme)
        res['종합점수'] = (res['시가총액'].rank(pct=True) * 50 + res['거래량'].rank(pct=True) * 50).round(2)
        res['시가총액(억)'] = (res['시가총액'] / 100000000).astype(int)

        # 5. 최종 정리 (상위 100개)
        final = res[['종합점수', '주도테마', '종목명', '현재가', 'PER', 'PBR', '거래량', '시가총액(억)']].sort_values('종합점수', ascending=False).head(100)

        # 6. 엑셀 저장 및 메일 발송
        fname = f"Stock_Report_{date}.xlsx"
        final.to_excel(fname, index=False)

        email_user = "chomiryo8462@gmail.com"
        email_pw = os.environ.get('EMAIL_PASSWORD')
        
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 주식 리포트 완료 ({date})"
        msg['To'] = email_user
        msg['From'] = email_user

        with open(fname, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={fname}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_pw)
            server.sendmail(email_user, email_user, msg.as_string())
        print("✅ 성공적으로 발송되었습니다.")

    except Exception as e:
        print(f"❌ 오류 발생: {e}")

if __name__ == "__main__":
    send_email()
