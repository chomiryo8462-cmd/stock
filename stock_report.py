import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from pykrx import stock

def get_valid_date():
    """데이터가 존재하는 가장 최근 영업일을 찾습니다."""
    curr = datetime.now()
    for _ in range(10):
        d_str = curr.strftime('%Y%m%d')
        df = stock.get_market_ohlcv_by_ticker(d_str, market="KOSPI")
        if not df.empty: return d_str
        curr -= timedelta(days=1)
    return datetime.now().strftime('%Y%m%d')

def send_email():
    try:
        target_date = get_valid_date()
        print(f"기준 날짜: {target_date}")

        # 1. 시세 및 지표 데이터 수집
        df_p = stock.get_market_ohlcv_by_ticker(target_date, market="ALL")
        df_f = stock.get_market_fundamental_by_ticker(target_date, market="ALL")
        
        # 2. 데이터 병합 (모든 컬럼을 일단 합침)
        df = pd.merge(df_p, df_f, left_index=True, right_index=True, how='outer')
        
        # 3. 종목명 추가 및 인덱스 정리
        tickers = df.index.tolist()
        df['종목명'] = [stock.get_market_ticker_name(t) for t in tickers]
        df = df.reset_index()

        # 4. 컬럼명 유연하게 매칭 (이름이 달라도 위치나 유사어로 찾음)
        def get_col(possible_list):
            for p in possible_list:
                for c in df.columns:
                    if p.lower() in str(c).lower(): return c
            return None

        c_close = get_col(['종가', 'Close'])
        c_vol = get_col(['거래량', 'Volume'])
        c_marcap = get_col(['시가총액', 'Marcap'])
        c_per = get_col(['PER'])
        c_pbr = get_col(['PBR'])
        c_div = get_col(['배당수익률', 'DIV'])
        c_chg = get_col(['등락률', 'ChgRate'])

        # 5. 수치 데이터 변환
        for c in [c_close, c_vol, c_marcap, c_per, c_pbr, c_div, c_chg]:
            if c: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

        # 6. 테마 분류 및 종합점수
        def theme(n):
            if any(k in n for k in ['삼성전자', 'SK하이닉스', '한미반도체']): return 'AI/반도체'
            if any(k in n for k in ['제룡전기', '효성중공업', '현대일렉트릭', '두산에너빌리티']): return '전력/에너지'
            if any(k in n for k in ['대한광통신', '오이솔루션']): return '광통신/5G'
            if any(k in n for k in ['YG', '하이브', '금융']): return '엔터/금융'
            return '기타'

        df['주도테마'] = df['종목명'].apply(theme)
        df['종합점수'] = (df[c_marcap].rank(pct=True) * 50 + df[c_vol].rank(pct=True) * 50).round(2)
        df['시가총액(억)'] = (df[c_marcap] / 100000000).astype(int)

        # 7. 최종 결과 구성
        res = pd.DataFrame({
            '종합점수': df['종합점수'],
            '주도테마': df['주도테마'],
            '종목명': df['종목명'],
            '현재가': df[c_close],
            '등락률(%)': df[c_chg],
            'PER': df[c_per],
            'PBR': df[c_pbr],
            '배당수익률(%)': df[c_div],
            '거래량': df[c_vol],
            '시가총액(억)': df['시가총액(억)']
        })

        df_final = res.sort_values(by='종합점수', ascending=False).head(100)

        # 8. 엑셀 저장 및 메일 발송
        fname = f"Stock_Report_{target_date}.xlsx"
        df_final.to_excel(fname, index=False)

        email_user = "chomiryo8462@gmail.com"
        email_pw = os.environ.get('EMAIL_PASSWORD')
        
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 주식 마스터 리포트 ({target_date})"
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
        print(f"✅ 전송 성공: {target_date}")

    except Exception as e:
        print(f"❌ 에러: {e}")

if __name__ == "__main__":
    send_email()
