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
        # 1. 데이터 수집
        df = fdr.StockListing('KRX')
        
        # 2. 항목 이름 자동 매칭 (대소문자 무관하게 찾기)
        # 로봇이 헷갈려하는 시가총액, 등락률 등을 자동으로 찾아줍니다.
        col_map = {c.lower(): c for c in df.columns}
        
        def get_col(name_list):
            for name in name_list:
                if name.lower() in col_map:
                    return col_map[name.lower()]
            return None

        target_chg = get_col(['ChgRate', 'ChangesRatio', '등락률'])
        target_marcap = get_col(['MarCap', 'Marcap', '시가총액'])
        target_vol = get_col(['Volume', '거래량'])

        # 3. 데이터 수치화 및 점수 계산
        # 항목이 없더라도 에러 없이 넘어가도록 설정했습니다.
        for col in [target_chg, target_marcap, target_vol]:
            if col:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 주말에도 순위가 나오도록 시가총액과 거래량 위주로 점수 산정
        df['종합점수'] = 0
        if target_marcap: df['종합점수'] += df[target_marcap].rank(pct=True) * 50
        if target_vol: df['종합점수'] += df[target_vol].rank(pct=True) * 50
        df['종합점수'] = df['종합점수'].round(2)

        # 4. 한글 이름표 붙이기
        rename_dict = {
            'Name': '종목명', 'Close': '현재가', 'Market': '시장'
        }
        if target_chg: rename_dict[target_chg] = '등락률(%)'
        if target_vol: rename_dict[target_vol] = '거래량'
        if target_marcap: rename_dict[target_marcap] = '시가총액(억)'
        
        df = df.rename(columns=rename_dict)

        # 5. 최종 리포트 구성 (존재하는 항목만 쏙쏙 골라 담기)
        cols = ['종합점수', '종목명', '현재가', '등락률(%)', '거래량', '시가총액(억)', '시장']
        valid_cols = [c for c in cols if c in df.columns]
        df_final = df[valid_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 6. 파일 저장 및 전송
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Report_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='openpyxl')

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')

        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [보안강화] 주식 종합 분석 리포트 ({today})"
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
        
        print("✅ 드디어 성공! 리포트가 발송되었습니다.")

    except Exception as e:
        print(f"❌ 최종 방어선 돌파됨(오류): {str(e)}")

if __name__ == "__main__":
    send_email()
