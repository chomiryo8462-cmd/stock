import pandas as pd
import FinanceDataReader as fdr
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import traceback

# =========================================
# 테마 정의
# =========================================

THEMES = {

    'AI/반도체': [
        '삼성전자',
        'SK하이닉스',
        '한미반도체',
        '리노공업'
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
        '디앤디파마텍'
    ]
}


# =========================================
# 테마 분류
# =========================================

def classify_theme(name):

    for theme, keywords in THEMES.items():

        if any(k in name for k in keywords):
            return theme

    return '기타'


# =========================================
# 최근 영업일
# =========================================

def get_last_business_day():

    today = datetime.now()

    while today.weekday() >= 5:
        today -= timedelta(days=1)

    return today.strftime("%Y-%m-%d")


# =========================================
# 한국 주식 리스트
# =========================================

def get_stock_list():

    kospi = fdr.StockListing('KOSPI')
    kosdaq = fdr.StockListing('KOSDAQ')

    stocks = pd.concat([kospi, kosdaq])

    return stocks[['Code', 'Name']]


# =========================================
# 시총 크롤링
# =========================================

def get_market_cap(code):

    try:

        url = f"https://finance.naver.com/item/main.naver?code={code}"

        headers = {
            "User-Agent": "Mozilla/5.0"
        }

        r = requests.get(
            url,
            headers=headers,
            timeout=10
        )

        soup = BeautifulSoup(r.text, 'lxml')

        em = soup.select_one(
            "#_market_sum"
        )

        if em is None:
            return 0

        text = em.text.strip()

        text = text.replace(",", "")

        return float(text)

    except:
        return 0


# =========================================
# 개별 종목 데이터
# =========================================

def get_stock_data(code, name):

    try:

        end = datetime.now()

        start = end - timedelta(days=30)

        df = fdr.DataReader(
            code,
            start.strftime("%Y-%m-%d"),
            end.strftime("%Y-%m-%d")
        )

        if df.empty:
            return None

        latest = df.iloc[-1]

        prev = df.iloc[-2]

        close = latest['Close']

        volume = latest['Volume']

        change = (
            (close - prev['Close'])
            / prev['Close']
        ) * 100

        value = close * volume

        market_cap = get_market_cap(code)

        return {

            '종목명': name,

            '현재가': round(close, 2),

            '등락률': round(change, 2),

            '거래량': int(volume),

            '거래대금': int(value),

            '시가총액(억)': market_cap,

            '주도테마': classify_theme(name)
        }

    except Exception as e:

        print(f"에러: {name} {e}")

        return None


# =========================================
# 데이터 수집
# =========================================

def collect_all_data():

    stocks = get_stock_list()

    results = []

    total = len(stocks)

    for idx, row in stocks.iterrows():

        code = row['Code']
        name = row['Name']

        print(f"[{idx+1}/{total}] {name}")

        data = get_stock_data(code, name)

        if data:
            results.append(data)

    df = pd.DataFrame(results)

    if df.empty:
        raise ValueError("수집 데이터 없음")

    # 점수 계산
    df['종합점수'] = (

        df['등락률'].rank(pct=True) * 40 +

        df['거래대금'].rank(pct=True) * 40 +

        df['시가총액(억)'].rank(pct=True) * 20

    ).round(2)

    df = df.sort_values(
        by='종합점수',
        ascending=False
    )

    return df.head(100)


# =========================================
# 엑셀 저장
# =========================================

def save_excel(df):

    now = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    filename = f"Stock_Report_{now}.xlsx"

    df.to_excel(
        filename,
        index=False
    )

    print(f"엑셀 저장 완료: {filename}")

    return filename


# =========================================
# 이메일 발송
# =========================================

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
        "📊 자동 주식 리포트"
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


# =========================================
# 메인
# =========================================

def main():

    try:

        print("데이터 수집 시작")

        df = collect_all_data()

        print(df.head())

        filename = save_excel(df)

        send_email(filename)

        print("전체 완료")

    except Exception as e:

        print("최종 에러")

        print(str(e))

        traceback.print_exc()


# =========================================

if __name__ == "__main__":
    main()
