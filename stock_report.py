import pandas as pd
import FinanceDataReader as fdr
from pykrx import stock
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import traceback
import os
import time

# =========================================================
# 설정
# =========================================================

TOP_N = 100
VOLUME_TOP = 300

THEMES = {

    'AI/반도체': [
        '삼성전자',
        'SK하이닉스',
        '한미반도체',
        '리노공업',
        '이오테크닉스'
    ],

    '전력/에너지': [
        '제룡전기',
        '효성중공업',
        '현대일렉트릭',
        '두산에너빌리티'
    ],

    '바이오': [
        '알테오젠',
        '리가켐바이오',
        '디앤디파마텍',
        '삼천당제약'
    ],

    '로봇': [
        '레인보우로보틱스',
        '두산로보틱스'
    ]
}


# =========================================================
# 최근 영업일
# =========================================================

def get_last_business_day():

    today = datetime.now()

    while today.weekday() >= 5:
        today -= timedelta(days=1)

    return today.strftime("%Y%m%d")


# =========================================================
# 테마 분류
# =========================================================

def classify_theme(name):

    for theme, keywords in THEMES.items():

        if any(k in name for k in keywords):
            return theme

    return '기타'


# =========================================================
# 뉴스 개수 수집
# =========================================================

def get_news_count(keyword):

    try:

        url = (
            f"https://search.naver.com/search.naver?"
            f"where=news&query={keyword}"
        )

        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        r = requests.get(
            url,
            headers=headers,
            timeout=10
        )

        soup = BeautifulSoup(r.text, 'lxml')

        news_items = soup.select(
            ".news_area"
        )

        return len(news_items)

    except:
        return 0


# =========================================================
# 거래량 상위 종목 추출
# =========================================================

def get_volume_top_stocks(target_date):

    df = stock.get_market_ohlcv_by_ticker(
        target_date,
        market="ALL"
    )

    if df.empty:
        raise ValueError("OHLCV 데이터 없음")

    df = df.sort_values(
        by='거래량',
        ascending=False
    )

    df = df.head(VOLUME_TOP)

    return df


# =========================================================
# 종목 상세 분석
# =========================================================

def analyze_stock(ticker, target_date):

    try:

        name = stock.get_market_ticker_name(
            ticker
        )

        # 가격 데이터
        price_df = stock.get_market_ohlcv_by_ticker(
            target_date,
            market="ALL"
        )

        row = price_df.loc[ticker]

        close = row['종가']
        volume = row['거래량']
        change = row['등락률']

        trading_value = close * volume

        # 펀더멘털
        fund_df = stock.get_market_fundamental_by_ticker(
            target_date,
            market="ALL"
        )

        if ticker in fund_df.index:

            fund = fund_df.loc[ticker]

            per = fund.get('PER', 0)
            pbr = fund.get('PBR', 0)
            div = fund.get('DIV', 0)

        else:

            per = 0
            pbr = 0
            div = 0

        # 뉴스
        news_count = get_news_count(name)

        # 테마
        theme = classify_theme(name)

        return {

            '종목명': name,

            '현재가': round(close, 2),

            '등락률': round(change, 2),

            '거래량': int(volume),

            '거래대금': int(trading_value),

            'PER': round(per, 2),

            'PBR': round(pbr, 2),

            '배당률': round(div, 2),

            '뉴스수': news_count,

            '주도테마': theme
        }

    except Exception as e:

        print(f"에러 발생 {ticker}: {e}")

        return None


# =========================================================
# 전체 분석
# =========================================================

def collect_data():

    target_date = get_last_business_day()

    print(f"기준일: {target_date}")

    volume_df = get_volume_top_stocks(
        target_date
    )

    results = []

    tickers = volume_df.index.tolist()

    total = len(tickers)

    for idx, ticker in enumerate(tickers):

        print(f"[{idx+1}/{total}] {ticker}")

        data = analyze_stock(
            ticker,
            target_date
        )

        if data:
            results.append(data)

        time.sleep(0.2)

    df = pd.DataFrame(results)

    if df.empty:
        raise ValueError("결과 없음")

    # =====================================================
    # 종합점수 계산
    # =====================================================

    df['종합점수'] = (

        df['등락률'].rank(pct=True) * 25 +

        df['거래량'].rank(pct=True) * 20 +

        df['거래대금'].rank(pct=True) * 20 +

        df['뉴스수'].rank(pct=True) * 20 +

        (1 / (df['PER'].replace(0, 9999))).rank(pct=True) * 10 +

        (1 / (df['PBR'].replace(0, 9999))).rank(pct=True) * 5

    ).round(2)

    df = df.sort_values(
        by='종합점수',
        ascending=False
    )

    return df.head(TOP_N)


# =========================================================
# 엑셀 저장
# =========================================================

def save_excel(df):

    now = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    filename = (
        f"Stock_Report_{now}.xlsx"
    )

    df.to_excel(
        filename,
        index=False
    )

    print(f"엑셀 저장 완료: {filename}")

    return filename


# =========================================================
# 이메일 전송
# =========================================================

def send_email(filename):

    email_user = os.environ.get(
        'EMAIL_USER'
    )

    email_pw = os.environ.get(
        'EMAIL_PASSWORD'
    )

    if not email_user:
        raise ValueError(
            "EMAIL_USER 없음"
        )

    if not email_pw:
        raise ValueError(
            "EMAIL_PASSWORD 없음"
        )

    msg = MIMEMultipart()

    msg['Subject'] = (
        "📊 오늘의 한국 주식 리포트"
    )

    msg['To'] = email_user
    msg['From'] = email_user

    with open(filename, "rb") as f:

        part = MIMEBase(
            "application",
            "octet-stream"
        )

        part.set_payload(f.read())

        encoders.encode_base64(part)

        part.add_header(
            "Content-Disposition",
            f"attachment; filename={filename}"
        )

        msg.attach(part)

    print("이메일 전송 시작")

    with smtplib.SMTP_SSL(
        'smtp.gmail.com',
        465
    ) as server:

        server.login(
            email_user,
            email_pw
        )

        server.sendmail(
            email_user,
            email_user,
            msg.as_string()
        )

    print("이메일 전송 완료")


# =========================================================
# 메인
# =========================================================

def main():

    try:

        print("주식 데이터 분석 시작")

        df = collect_data()

        print(df.head())

        filename = save_excel(df)

        send_email(filename)

        print("전체 작업 완료")

    except Exception as e:

        print("최종 에러")

        print(str(e))

        traceback.print_exc()


# =========================================================

if __name__ == "__main__":
    main()
