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
        
        # 2. 항목 찾기 및 숫자 변환
        col_map = {c.lower(): c for c in df.columns}
        def get_col(names):
            for n in names:
                if n.lower() in col_map: return col_map[n.lower()]
            return None

        t_chg = get_col(['ChgRate', 'ChangesRatio', '등락률'])
        t_marcap = get_col(['MarCap', 'Marcap', '시가총액'])
        t_vol = get_col(['Volume', '거래량'])

        for c in [t_chg, t_marcap, t_vol]:
            if c: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

        # 3. 종합점수 계산 (시가총액 50 + 거래량 50)
        df['종합점수'] = 0
        if t_marcap: df['종합점수'] += df[t_marcap].rank(pct=True) * 50
        if t_vol: df['종합점수'] += df[t_vol].rank(pct=True) * 50
        df['종합점수'] = df['종합점수'].round(2)

        # 4. 시가총액 단위 조정 (지수 표현 방지)
        # 원본 데이터는 원 단위이므로 1억(100,000,000)으로 나눠 억 단위로 표기
        if t_marcap:
            df['시가총액(억)'] = (df[t_marcap] / 100000000).astype(int)

        # 5. 한글 이름표 및 정렬
        rename_dict = {'Name': '종목명', 'Close': '현재가', 'Market': '시장'}
        if t_chg: rename_dict[t_chg] = '등락률(%)'
        if t_vol: rename_dict[t_vol] = '거래량'
        
        df = df.rename(columns=rename_dict)
        final_cols = ['종합점수', '종목명', '현재가', '등락률(%)', '거래량', '시가총액(억)', '시장']
        df_final = df[[c for c in final_cols if c in df.columns]].sort_values(by='종합점수', ascending=False).head(100)

        # 6. 엑셀 저장
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Report_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='openpyxl')

        # 7. 메일 전송
        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        msg = MIMEMultipart()
        msg['Subject'] = f"📈 [숫자개선] 주식 종합 분석 리포트 ({today})"
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
        
        print("✅ 가독성 개선 리포트 발송 성공!")

    except Exception as e:
        print(f"❌ 오류: {str(e)}")

if __name__ == "__main__":
    send_email()
