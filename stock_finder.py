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
        # 1. 기본 종목 리스트 (KRX 전체)
        df_base = fdr.StockListing('KRX')
        
        # 2. 투자지표(PER/PBR/배당) 데이터 확보
        # FinanceDataReader의 최신 버전에서는 KRX-DESC 등을 통해 상세 지표를 가져옵니다.
        # 데이터 유실 방지를 위해 컬럼명을 대소문자 구분 없이 매칭합니다.
        col_map = {c.lower(): c for c in df_base.columns}
        
        def find_col(possible_names):
            for name in possible_names:
                if name.lower() in col_map:
                    return col_map[name.lower()]
            return None

        # 필수 항목 매칭
        t_marcap = find_col(['MarCap', 'Marcap', '시가총액'])
        t_vol = find_col(['Volume', '거래량'])
        t_chg = find_col(['ChgRate', 'ChangesRatio', '등락률'])
        t_per = find_col(['PER'])
        t_pbr = find_col(['PBR'])
        t_div = find_col(['DividendYield', '배당수익률'])

        # 3. 데이터 숫자 변환 및 전처리
        process_cols = [t_marcap, t_vol, t_chg, t_per, t_pbr, t_div, 'Close']
        for c in process_cols:
            if c:
                df_base[c] = pd.to_numeric(df_base[c], errors='coerce').fillna(0)

        # 4. 주도 테마 분류 로직 (AI, 반도체, 전력, 광통신)
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
        
        df_base['주도테마'] = df_base['Name'].apply(classify_theme)

        # 5. 종합점수 계산 (시총 50 + 거래량 50)
        df_base['종합점수'] = 0
        if t_marcap: df_base['종합점수'] += df_base[t_marcap].rank(pct=True) * 50
        if t_vol: df_base['종합점수'] += df_base[t_vol].rank(pct=True) * 50
        df_base['종합점수'] = df_base['종합점수'].round(2)

        # 6. 리포트 항목 정리 및 한글화
        if t_marcap:
            df_base['시가총액(억)'] = (df_base[t_marcap] / 100000000).astype(int)
        
        # 엑셀에 표시할 최종 데이터프레임 구성
        rename_map = {
            'Name': '종목명', 'Close': '현재가', t_chg: '등락률(%)',
            t_per: 'PER', t_pbr: 'PBR', t_div: '배당수익률(%)',
            t_vol: '거래량', 'Market': '시장'
        }
        df_renamed = df_base.rename(columns={k: v for k, v in rename_map.items() if k})

        # 컬럼이 없을 경우 빈 값으로라도 생성
        final_headers = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', 'PER', 'PBR', '배당수익률(%)', '거래량', '시가총액(억)', '시장']
        for header in final_headers:
            if header not in df_renamed.columns:
                df_renamed[header] = 0

        df_final = df_renamed[final_headers].sort_values(by='종합점수', ascending=False).head(100)

        # 7. 엑셀 저장 및 메일 발송
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Master_Report_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='openpyxl')

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [마스터리포트] 가치지표 및 테마 분석 ({today})"
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
        
        print("✅ 가치 지표 포함 리포트 전송 성공!")

    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    send_email()
