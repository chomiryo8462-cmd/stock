import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from pykrx import stock
import traceback

# =========================
# 테마 정의
# =========================
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
        '삼천당제약',
        '알테오젠',
        '리가켐바이오',
        '디앤디파마텍'
    ]
}


# =========================
# 최근 영업일 찾기
# =========================
def get_last_open_date():

    curr = datetime.now()

    for _ in range(10):

        d_str = curr.strftime('%Y%m%d')

        try:
            df = stock.get_market_ohlcv_by_ticker(
                d_str,
                market="KOSPI"
            )

            if not df.empty:
                return d_str

        except:
            pass

        curr -= timedelta(days=1)

    return datetime.now().strftime('%Y%m%d')


# =========================
# 테마 분류
# =========================
def classify_theme(name):

    for theme, keywords in THEMES.items():

        if any(k in name for k in keywords):
            return theme

    return '기타'


# =========================
# 데이터 수집
# =========================
def collect_stock_data(target_date):

    print(f"📅 데이터 기준일: {target_date}")

    # 시세
    df_price = stock.get_market_ohlcv_by_ticker(
        target_date,
        market="ALL"
    )

    # 펀더멘털
    df_fund = stock.get_market_fundamental_by_ticker(
        target_date,
        market="ALL"
    )

    # 병합
    df = pd.concat([df_price, df_fund], axis=1)

    # 종목명 추가
    df['종목명'] = [
        stock.get_market_ticker_name(t)
        for t in df.index
    ]

    # 필요한 컬럼만 안정적으로 추출
    result = pd.DataFrame(index=df.index)

    result['종목명'] = df['종목명']

    result['현재가'] = df.get('종가', 0)
    result['등락률'] = df.get('등락률', 0)
    result['거래량'] = df.get('거래량', 0)
    result['시가총액'] = df.get('시가총액', 0)

    # 기본지표
    result['PER'] = df.get('PER', 0)
    result['PBR'] = df.get('PBR', 0)
    result['DIV'] = df.get('DIV', 0)

    # NaN 방어
    result = result.fillna(0)

    # 거래대금
    result['거래대금'] = (
        result['현재가'] * result['거래량']
    )

    # 테마
    result['주도테마'] = result['종목명'].apply(
        classify_theme
    )

    # 시총 억단위
    result['시가총액(억)'] = (
        result['시가총액'] / 1e8
    ).round(1)

    # =========================
    # 종합점수 계산
    # =========================

    result['종합점수'] = (

        result['등락률'].rank(pct=True) * 35 +

        result['거래대금'].rank(pct=True) * 35 +

        result['시가총액'].rank(pct=True) * 20 +

        result['PER'].rank(
            pct=True,
            ascending=False
        ) * 10

    ).round(2)

    # 정렬
    final = result[
        [
            '종합점수',
            '주도테마',
            '종목명',
            '현재가',
            '등락률',
            '거래량',
            '거래대금',
            'PER',
            'PBR',
            '시가총액(억)'
        ]
    ]

    final = final.sort_values(
        by='종합점수',
        ascending=False
    )

    return final.head(100)


# =========================
# 엑셀 저장
# =========================
def save_excel(df, target_date):

    now = datetime.now().strftime('%H%M%S')

    filename = (
        f"Stock_Report_{target_date}_{now}.xlsx"
    )

    df.to_excel(
        filename,
        index=False
    )

    print(f"💾 엑셀 저장 완료: {filename}")

    return filename


# =========================
# 이메일 발송
# =========================
def send_email(filename, target_date):

    email_user = os.environ.get('EMAIL_USER')
    email_pw = os.environ.get('EMAIL_PASSWORD')

    if not email_user:
        raise ValueError(
            "EMAIL_USER 환경변수 없음"
        )

    if not email_pw:
        raise ValueError(
            "EMAIL_PASSWORD 환경변수 없음"
        )

    msg = MIMEMultipart()

    msg['Subject'] = (
        f"📊 주식 리포트 ({target_date})"
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

    print("📧 이메일 전송 시작...")

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

    print("🚀 이메일 전송 완료")


# =========================
# 메인 실행
# =========================
def main():

    try:

        target_date = get_last_open_date()

        df = collect_stock_data(target_date)

        filename = save_excel(
            df,
            target_date
        )

        send_email(
            filename,
            target_date
        )

        print("✅ 전체 작업 완료")

    except Exception as e:

        print("❌ 에러 발생")

        print(str(e))

        traceback.print_exc()


# =========================
# 시작
# =========================
if __name__ == "__main__":
    main()
