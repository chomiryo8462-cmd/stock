import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from pykrx import stock

def get_last_open_date():
    """데이터가 확실히 존재하는 가장 최근 영업일을 찾습니다."""
    curr = datetime.now()
    for _ in range(10): # 최대 10일 전까지 거슬러 올라감
        d_str = curr.strftime('%Y%m%d')
        df = stock.get_market_ohlcv_by_ticker(d_str, market="KOSPI")
        if not df.empty: return d_str
        curr -= timedelta(days=1)
    return datetime.now().strftime('%Y%m%d')

def send_email():
    try:
        target_date = get_last_open_date()
        print(f"✅ 기준 날짜 확인: {target_date}")

        # 1. 데이터 수집 (시세 및 기본지표)
        df_p = stock.get_market_ohlcv_by_ticker(target_date, market="ALL")
        df_f = stock.get_market_fundamental_by_ticker(target_date, market="ALL")
        
        # 2. 데이터 병합
        df = pd.concat([df_p, df_f], axis=1)
        df['종목명'] = [stock.get_market_ticker_name(t) for t in df.index]
        
        # 3. [핵심] 이름 대신 '위치'로 데이터 추출 (에러 원천 차단)
        res = pd.DataFrame(index=df.index)
        res['종목명'] = df['종목명']
        res['현재가'] = df.iloc[:, 3]    # 4번째 열: 종가
        res['등락률'] = df.iloc[:, 5]    # 6번째 열: 등락률
        res['거래량'] = df.iloc[:, 6]    # 7번째 열: 거래량
        res['시가총액'] = df.iloc[:, 7]  # 8번째 열: 시가총액
        
        # 지표 데이터 (PER, PBR 등은 이름으로 찾되 없으면 0)
        for col in ['PER', 'PBR', 'DVD_YLD']:
            res[col] = df[col] if col in df.columns else 0

        # 4. 테마 분류 및 점수 계산
        def classify(n):
            if any(k in n for k in ['삼성전자', 'SK하이닉스', '한미반도체']): return 'AI/반도체'
            if any(k in n for k in ['제룡전기', '효성중공업', '현대일렉트릭', '두산에너빌리티']): return '전력/에너지'
            return '기타'

        res['주도테마'] = res['종목명'].apply(classify)
        res['종합점수'] = (res['시가총액'].rank(pct=True) * 50 + res['거래량'].rank(pct=True) * 50).round(2)
        res['시가총액(억)'] = (res['시가총액'] / 100000000).fillna(0).astype(int)

        # 5. 상위 100개 정리
        final = res[['종합점수', '주도테마', '종목명', '현재가', '등락률', 'PER', 'PBR', '시가총액(억)']]
        final = final.sort_values(by='종합점수', ascending=False).head(100)

        # 6. 엑셀 저장 및 발송
        fname = f"Stock_Report_{target_date}.xlsx"
        final.to_excel(fname, index=False)

        email_user = "chomiryo8462@gmail.com"
        email_pw = os.environ.get('EMAIL_PASSWORD')
        
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 주식 리포트 발송 ({target_date})"
        msg['To'] = email_user
        msg['From'] = email_user

        with open(fname, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read()); encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={fname}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_pw)
            server.sendmail(email_user, email_user, msg.as_string())
        print("🚀 전송 완료!")

    except Exception as e:
        print(f"❌ 최종 에러 발생: {e}")

if __name__ == "__main__":
    send_email()
