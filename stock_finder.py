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
        # 1. 통합 데이터 및 상세 지표(PER, PBR 등) 가져오기
        # KRX 전체 종목 리스트와 상세 지표를 병합합니다.
        df_base = fdr.StockListing('KRX')
        
        # 2. 항목 이름 자동 매칭 및 숫자 변환
        col_map = {c.lower(): c for c in df_base.columns}
        def get_col(names):
            for n in names:
                if n.lower() in col_map: return col_map[n.lower()]
            return None

        t_chg = get_col(['ChgRate', 'ChangesRatio', '등락률'])
        t_marcap = get_col(['MarCap', 'Marcap', '시가총액'])
        t_vol = get_col(['Volume', '거래량'])

        # 숫자 변환
        for c in [t_chg, t_marcap, t_vol, 'Close']:
            if c: df_base[c] = pd.to_numeric(df_base[c], errors='coerce').fillna(0)

        # 3. 대세 테마(AI, 반도체, 전력, 광통신 등) 키워드 매칭 로직
        def detect_theme(name):
            themes = {
                'AI/반도체': ['삼성전자', 'SK하이닉스', '한미반도체', '리노공업', 'HBM', 'AI'],
                '전력/에너지': ['제룡전기', '효성중공업', 'LS', 'HD현대일렉트릭', '두산에너빌리티', '변압기'],
                '광통신/통신': ['대한광통신', '오이솔루션', '서진시스템', '광통신', '5G'],
                '엔터/문화': ['YG', '하이브', 'JYP', '에스엠']
            }
            for theme, keywords in themes.items():
                if any(key in name for key in keywords):
                    return theme
            return '기타'

        df_base['주도테마'] = df_base['Name'].apply(detect_theme)

        # 4. 종합점수 계산 (시총 50 + 거래량 50)
        df_base['종합점수'] = 0
        if t_marcap: df_base['종합점수'] += df_base[t_marcap].rank(pct=True) * 50
        if t_vol: df_base['종합점수'] += df_base[t_vol].rank(pct=True) * 50
        df_base['종합점수'] = df_base['종합점수'].round(2)

        # 5. 단위 조정 및 항목 정리
        if t_marcap:
            df_base['시가총액(억)'] = (df_base[t_marcap] / 100000000).astype(int)
        
        rename_dict = {
            'Name': '종목명', 'Close': '현재가', t_chg: '등락률(%)', 
            t_vol: '거래량', 'Market': '시장'
        }
        df_base = df_base.rename(columns=rename_dict)

        # 6. 최종 리포트 구성 (상위 100개)
        # 참고: PER, PBR, 배당수익률은 데이터 소스에 따라 추가 연산이 필요할 수 있으나 
        # 현재는 기본 제공되는 등락률과 시총 위주로 구성하며 테마를 강조했습니다.
        final_cols = ['종합점수', '주도테마', '종목명', '현재가', '등락률(%)', '거래량', '시가총액(억)', '시장']
        df_final = df_base[final_cols].sort_values(by='종합점수', ascending=False).head(100)

        # 7. 엑셀 저장 및 발송
        today = datetime.now().strftime('%Y-%m-%d')
        filename = f"Stock_Market_Leader_Report_{today}.xlsx"
        df_final.to_excel(filename, index=False, engine='openpyxl')

        email_user = "chomiryo8462@gmail.com"
        email_password = os.environ.get('EMAIL_PASSWORD')
        msg = MIMEMultipart()
        msg['Subject'] = f"🔥 [주도주분석] AI/전력/광통신 리포트 ({today})"
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
        
        print("✅ 테마별 주도주 리포트 발송 완료!")

    except Exception as e:
        print(f"❌ 오류: {str(e)}")

if __name__ == "__main__":
    send_email()
