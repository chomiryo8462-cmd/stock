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
        # 1. 데이터 수집 (안전하게 항목 찾기)
        df = fdr.StockListing('KRX')
        
        # 2. 대소문자 무관하게 항목 매칭 (MarCap, Marcap 모두 대응)
        col_map = {c.lower(): c for c in df.columns}
        def get_actual_col(standard_name):
            return col_map.get(standard_name.lower())

        target_chg = get_actual_col('ChgRate') or get_actual_col('ChangesRatio')
        target_marcap = get_actual_col('MarCap') or get_actual_col('Marcap')
        target_vol = get_actual_col('Volume')

        # 3. 데이터 숫자 변환
        check_cols = [target_chg, target_marcap, target_vol, 'Close', 'PER', 'PBR', 'DividendYield']
        for c in check_cols:
            if c and c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

        # 4. 주도 테마 분류 (AI/반도체, 전력, 광통신)
        def classify_theme(name):
            themes = {
                'AI/반도체': ['삼성전자', 'SK하이닉스', '한미반도체', 'HBM', 'AI'],
                '전력/에너지': ['제룡전기', '효성중공업', 'HD현대일렉트릭', '두산에너빌리티', '변압기'],
                '광통신/5G': ['대한광통신', '오이솔루션', '광통신', '5G'],
                '엔터/금융': ['YG', '하이브', 'KB금융', '신한지주']
            }
            for theme, keywords in themes.items():
                if any(key in name for key in keywords): return theme
            return '기타'
        
        df['주도테마'] = df['Name'].apply(classify_theme)

        # 5. 종합점수 계산 (시총 50 + 거래량 50)
        df['종합점수'] = 0
        if target_marcap: df['종합점수'] += df[target_marcap].rank(pct=True) * 50
        if target_vol: df['종합점수'] += df[target_vol].rank(pct=True) * 50
        df['종합점수'] = df['종합점수'].round(2)

        # 6. 항목 한글화 및 단위 조정
        if target_marcap:
            df['시가총액(억)'] = (df[target_marcap] / 100000000).astype(int)
        
        rename_dict = {
            'Name': '종목명', 'Close': '현재가', target_chg: '등락률(%)',
            'PER': 'PER', 'PBR': 'PBR', 'DividendYield': '배당수익률(%)',
            target_vol: '거래량', 'Market': '시장'
        }
        df = df.rename(columns=rename_dict)

        # 7. 최종 리포트 구성 (존재하는 항목만 선별)
        desired = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액(억)', '시장']
        final_cols = [c for c in desired if c in df.columns]
        df_final = df[final_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 8. 파일 저장 및 전송
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Master_Report_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='openpyxl')

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        msg = MIMEMultipart()
        msg['Subject'] = f"🏆 [마스터] AI/전력/가치지표 종합 분석 ({today})"
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
        
        print("✅ 마스터 리포트 전송 성공!")

    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    send_email()
