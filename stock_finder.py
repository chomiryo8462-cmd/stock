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
        # 1. 데이터 수집 (KRX 전체)
        df_base = fdr.StockListing('KRX')
        
        # 2. 컬럼 매칭 및 숫자 변환 (PER, PBR 0값 방지 시도)
        col_map = {c.lower(): c for c in df_base.columns}
        def find_col(possible_names):
            for name in possible_names:
                if name.lower() in col_map: return col_map[name.lower()]
            return None

        t_per, t_pbr, t_div = find_col(['per']), find_col(['pbr']), find_col(['dividendyield', '배당수익률'])
        t_marcap, t_vol, t_chg = find_col(['marcap', '시가총액']), find_col(['volume', '거래량']), find_col(['chgrate', '등락률'])

        for c in [t_marcap, t_vol, t_chg, t_per, t_pbr, t_div, 'Close']:
            if c: df_base[c] = pd.to_numeric(df_base[c], errors='coerce').fillna(0)

        # 3. 테마 분류 및 종합점수 계산
        def classify_theme(name):
            themes = {'AI/반도체': ['삼성전자', 'SK하이닉스', '한미반도체'], '전력/에너지': ['제룡전기', '두산에너빌리티'], '광통신': ['대한광통신'], '엔터/금융': ['YG', 'KB금융']}
            for theme, keys in themes.items():
                if any(k in name for k in keys): return theme
            return '기타'
        
        df_base['주도테마'] = df_base['Name'].apply(classify_theme)
        df_base['종합점수'] = (df_base[t_marcap].rank(pct=True) * 50 + df_base[t_vol].rank(pct=True) * 50).round(2)
        df_base['시가총액(억)'] = (df_base[t_marcap] / 100000000).astype(int)

        # 4. 리포트 정리 (상위 100개)
        final_headers = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액(억)', '시장']
        df_final = df_base.rename(columns={'Name':'종목명', 'Close':'현재가', t_chg:'등락률(%)', t_per:'PER', t_pbr:'PBR', t_div:'배당수익률(%)', t_vol:'거래량', 'Market':'시장'})[final_headers]
        df_final = df_final.sort_values(by='종합점수', ascending=False).head(100)

        # 5. 엑셀 저장 및 메일 발송
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Report_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='openpyxl')

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [자동발송] 주식 마스터 리포트 ({today})"
        msg.attach(MIMEBase("application", "octet-stream")) # 파일 첨부 로직 생략(위 코드와 동일)
        # ... (생략된 메일 발송 코드 부분은 위와 동일하게 적용됨) ...
        
        with open(filename, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read()); encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_user, email_password)
            server.sendmail(email_user, email_user, msg.as_string())
        print("전송 완료!")
    except Exception as e: print(f"오류: {e}")

if __name__ == "__main__": send_email()
