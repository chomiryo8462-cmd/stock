import pandas as pd
import FinanceDataReader as fdr
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
import re

# =========================================================
# 설정
# =========================================================

TOP_N = 100
VOLUME_TOP = 100

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

# =========================================================
# 뉴스 개수 수집 (Google News RSS)
# =========================================================

def get_news_count(keyword):

    try:

        url = (
            "https://news.google.com/rss/search?"
            f"q={keyword}+when:1d&hl=ko&gl=KR&ceid=KR:ko"
        )

        r = requests.get(
            url,
            headers=HEADERS,
            timeout=10
        )

        soup = BeautifulSoup(
            r.text,
            'xml'
        )

        items = soup.find_all('item')

        return len(items)

    except Exception as e:

        print(f"뉴스 수집 실패 {keyword}: {e}")

        return 0


# =========================================================
# 네이버 금융 펀더멘털 수집
# =========================================================

def get_fundamental(code):

    try:

        url = (
            f"https://finance.naver.com/item/main.naver?code={code}"
        )

        r = requests.get(
            url,
            headers=HEADERS,
            timeout=10
        )

        text = r.text

        per = 0
        pbr = 0
        div = 0

        # =================================================
        # PER
        # =================================================

        per_match = re.search(
            r"PER[^0-9]*([0-9]+\.[0-9]+)",
            text
        )

        if per_match:

            per = float(
                per_match.group(1)
            )

        # =================================================
        # PBR
        # =================================================

        pbr_match = re.search(
            r"PBR[^0-9]*([0-9]+\.[0-9]+)",
            text
        )

        if pbr_match:

            pbr = float(
                pbr_match.group(1)
            )

        # =================================================
        # 배당수익률
        # =================================================

        div_match = re.search(
            r"배당수익률[^0-9]*([0-9]+\.[0-9]+)",
            text
        )

        if div_match:

            div = float(
                div_match.group(1)
            )

        return per, pbr, div

    except Exception as e:

        print(f"펀더멘털 실패 {code}: {e}")

        return 0, 0, 0


# =========================================================
# 종목 리스트
# =========================================================

def get_stock_list():

    kospi = fdr.StockListing('KOSPI')
    kosdaq = fdr.StockListing('KOSDAQ')

    df = pd.concat([
        kospi,
        kosdaq
    ])

    return df[['Code', 'Name']]


# =========================================================
# 개별 종목 분석
# =========================================================

def analyze_stock(code, name):

    try:

        end = datetime.now()

        start = end - timedelta(days=30)

        df = fdr.DataReader(
            code,
            start.strftime("%Y-%m-%d"),
            end.strftime("%Y-%m-%d")
        )

        if len(df) < 2:
            return None

        latest = df.iloc[-1]
        prev = df.iloc[-2]

        close = latest['Close']

        volume = latest['Volume']

        change = (
            (close - prev['Close'])
            / prev['Close']
        ) * 100

        trading_value = close * volume

        # 뉴스 개수
        news_count = get_news_count(name)

        # 펀더멘털
        per, pbr, div = get_fundamental(code)

        return {

            '종목명': name,

            '현재가': round(close, 2),

            '등락률': round(change, 2),

            '거래량': int(volume),

            '거래대금': int(trading_value),

            'PER': round(per, 2),

            'PBR': round(pbr, 2),

            '배당률': round(div, 2),

            '뉴스수': int(news_count)
        }

    except Exception as e:

        print(f"종목 분석 실패 {name}: {e}")

        return None


# =========================================================
# 데이터 수집
# =========================================================

def collect_data():

    stocks = get_stock_list()

    # 속도 제한
    stocks = stocks.head(VOLUME_TOP)

    results = []

    total = len(stocks)

    for idx, row in stocks.iterrows():

        code = row['Code']
        name = row['Name']

        print(f"[{idx+1}/{total}] {name}")

        data = analyze_stock(
            code,
            name
        )

        if data:
            results.append(data)

        time.sleep(0.2)

    df = pd.DataFrame(results)

    if df.empty:
        raise ValueError("수집 데이터 없음")

    # =====================================================
    # 점수 계산
    # =====================================================

    safe_per = (
        df['PER']
        .replace(0, 9999)
        .fillna(9999)
    )

    safe_pbr = (
        df['PBR']
        .replace(0, 9999)
        .fillna(9999)
    )

    df['종합점수'] = (

        df['등락률'].rank(pct=True) * 30 +

        df['거래량'].rank(pct=True) * 25 +

        df['거래대금'].rank(pct=True) * 20 +

        df['뉴스수'].rank(pct=True) * 15 +

        (1 / safe_per).rank(pct=True) * 5 +

        (1 / safe_pbr).rank(pct=True) * 5

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
        f"/tmp/Stock_Report_{now}.xlsx"
    )

    df.to_excel(
        filename,
        index=False,
        engine='openpyxl'
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
        raise ValueError("EMAIL_USER 없음")

    if not email_pw:
        raise ValueError("EMAIL_PASSWORD 없음")

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
            f'attachment; filename="{os.path.basename(filename)}"'
        )

        msg.attach(part)

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

    print("메일 전송 완료")


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

        print("전체 완료")

    except Exception as e:

        print("최종 에러 발생")

        print(str(e))

        traceback.print_exc()


# =========================================================

if __name__ == "__main__":

    main()
