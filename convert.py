"""
엑셀 → JSON 변환 스크립트
매출 DB_2026 시트를 읽어서 dashboard_data.json 생성
"""
import pandas as pd
import json
import sys
from pathlib import Path

EXCEL_PATH = Path("data/실적데이터.xlsx")
OUTPUT_PATH = Path("dashboard_data.json")

def safe(v):
    """NaN/inf → None 변환"""
    if pd.isna(v) or v != v:
        return None
    if isinstance(v, float) and (v == float('inf') or v == float('-inf')):
        return None
    return round(float(v), 4) if isinstance(v, float) else v

def load_db(xl, sheet, header_row):
    df = pd.read_excel(xl, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def monthly_sales(df, month_col='마감월', sales_col='원화판매금액'):
    """월별 매출 합계 (백만원)"""
    result = {}
    for m, grp in df.groupby(month_col):
        result[int(m)] = round(grp[sales_col].sum() / 1e6, 2)
    return result

def monthly_by_group(df, group_col, month_col='마감월', sales_col='원화판매금액'):
    """월별 × 그룹별 매출"""
    result = {}
    for (m, g), grp in df.groupby([month_col, group_col]):
        key = str(g)
        if key not in result:
            result[key] = {}
        result[key][int(m)] = round(grp[sales_col].sum() / 1e6, 2)
    return result

def brand_category(line):
    """라인명 → 브랜드 분류"""
    HERO_LINES = ['펩타이드 9', '피디알엔', '레티놀 콜라겐', '피토 이엑스 피디알엔', '영시카 피디알엔']
    CHAMP_LINES = ['레드 락토 콜라겐', '멜라논 엑스', '프리미엄 콜라겐', '그린 시카 콜라겐',
                   '엑스트라 슈퍼 9 플러스', 'EGF', '비건 비타민', '시카놀 B5']
    DAISO_CH = ['다이소']

    if line in HERO_LINES:
        return 'hero'
    if line in CHAMP_LINES:
        return 'champion'
    return 'other'

def channel_normalize(ch):
    """채널그룹 → 대분류"""
    if not ch or str(ch) == 'nan':
        return '기타'
    ch = str(ch)
    mapping = {
        'CIS권-B2B': 'CIS권', 'CIS권-B2C': 'CIS권',
        '동남아권-B2B': '동남아권', '동남아권-B2C': '동남아권',
        '중국권-B2B': '중국권', '중국권-B2C': '중국권',
        '일본권-B2B': '일본권', '일본권-B2C': '일본권',
        '북미권-B2B': '북미권', '북미권-B2C': '북미권',
        '동유럽권-B2B': '동유럽권', '동유럽권-B2C': '동유럽권',
        'Global권-B2B': 'Global권', 'Global권-B2C': 'Global권',
        '중동권-B2B': '중동권', '중동권-B2C': '중동권',
        '기타유럽권-B2B': '기타유럽권', '기타유럽권-B2C': '기타유럽권',
    }
    return mapping.get(ch, ch)

def convert():
    if not EXCEL_PATH.exists():
        print(f"ERROR: {EXCEL_PATH} 파일이 없습니다.")
        print("data/ 폴더에 실적데이터.xlsx 파일을 넣어주세요.")
        sys.exit(1)

    xl = pd.ExcelFile(EXCEL_PATH)
    print(f"파일 로드 완료: {EXCEL_PATH}")

    # ── 2026년 DB ──
    df26 = load_db(xl, '매출 DB_2026', 0)
    df26 = df26[df26['원화판매금액'].notna()]
    df26['채널_대분류'] = df26['채널그룹'].apply(channel_normalize)
    df26['브랜드'] = df26['라인'].apply(brand_category)
    df26['다이소여부'] = df26['유통구조'].apply(lambda x: '다이소' in str(x))

    # ── 2025년 DB ──
    df25 = load_db(xl, '매출 DB_2025', 4)
    df25 = df25[df25['추정 매출액'].notna()]
    # 25년 마감월: '25-1' → 1
    df25['마감월_int'] = df25['마감월'].apply(
        lambda x: int(str(x).replace('25-', '').strip()) if pd.notna(x) and str(x).startswith('25-') else None
    )
    df25 = df25[df25['마감월_int'].notna()]
    df25['채널_대분류'] = df25['유통구조'].apply(channel_normalize)
    def brand25(row):
        라인 = str(row.get('라인명', ''))
        구분2 = str(row.get('구분2', ''))
        HERO = ['펩타이드 9', '피디알엔', '레티날', '레티놀', '피토 이엑스 피디알엔']
        CHAMP = ['레드 락토 콜라겐', '멜라논', '프리미엄 콜라겐', '그린 시카 콜라겐',
                 '엑스트라 슈퍼 9', 'EGF', '비건 비타민', '시카놀']
        if any(h in 라인 for h in HERO): return 'hero'
        if any(c in 라인 for c in CHAMP): return 'champion'
        return 'other'
    df25['브랜드'] = df25.apply(brand25, axis=1)
    df25['다이소여부'] = df25['구분2'].apply(lambda x: '다이소' in str(x))

    # ════════════════════
    # 전사 월별 매출
    # ════════════════════
    monthly26 = monthly_sales(df26)
    monthly25_raw = df25.groupby('마감월_int')['추정 매출액'].sum() / 1e6
    monthly25 = {int(k): round(float(v), 2) for k, v in monthly25_raw.items()}

    # GP
    monthly_gp26 = {int(m): round(grp['이익액'].sum() / 1e6, 2)
                    for m, grp in df26.groupby('마감월')}
    monthly_gp25 = {int(m): round(grp['추정 매출원가'].sum() / 1e6, 2)  # 원가만, GP = 매출-원가
                    for m, grp in df25.groupby('마감월_int')}
    # 25년 GP 재계산
    monthly_gp25_real = {}
    for m, grp in df25.groupby('마감월_int'):
        sales = grp['추정 매출액'].sum()
        cost  = grp['추정 매출원가'].sum()
        monthly_gp25_real[int(m)] = round((sales - cost) / 1e6, 2)

    # ════════════════════
    # 국내/해외 분리
    # ════════════════════
    domestic26 = {int(m): round(grp['원화판매금액'].sum() / 1e6, 2)
                  for m, grp in df26[df26['국내/해외'] == '국내'].groupby('마감월')}
    overseas26 = {int(m): round(grp['원화판매금액'].sum() / 1e6, 2)
                  for m, grp in df26[df26['국내/해외'] == '해외'].groupby('마감월')}
    domestic25 = {}
    overseas25 = {}
    for m, grp in df25.groupby('마감월_int'):
        d = grp[grp['지역'] == '국내']['추정 매출액'].sum()
        o = grp[grp['지역'] != '국내']['추정 매출액'].sum()
        domestic25[int(m)] = round(d / 1e6, 2)
        overseas25[int(m)] = round(o / 1e6, 2)

    # ════════════════════
    # 브랜드별 월별 매출
    # ════════════════════
    brand26 = {}
    for (m, b), grp in df26.groupby(['마감월', '브랜드']):
        brand26.setdefault(b, {})[int(m)] = round(grp['원화판매금액'].sum() / 1e6, 2)

    brand25 = {}
    for (m, b), grp in df25.groupby(['마감월_int', '브랜드']):
        brand25.setdefault(b, {})[int(m)] = round(grp['추정 매출액'].sum() / 1e6, 2)

    # 다이소 별도
    daiso26 = {int(m): round(grp['원화판매금액'].sum() / 1e6, 2)
               for m, grp in df26[df26['다이소여부']].groupby('마감월')}
    daiso25 = {int(m): round(grp['추정 매출액'].sum() / 1e6, 2)
               for m, grp in df25[df25['다이소여부']].groupby('마감월_int')}

    # ════════════════════
    # 채널별 매출 (26년 누적)
    # ════════════════════
    ch26 = df26.groupby('채널_대분류')['원화판매금액'].sum().sort_values(ascending=False) / 1e6
    ch25_raw = df25.groupby('채널_대분류')['추정 매출액'].sum().sort_values(ascending=False) / 1e6
    channel_acc26 = {str(k): round(float(v), 2) for k, v in ch26.items()}
    channel_acc25 = {str(k): round(float(v), 2) for k, v in ch25_raw.items()}

    # ════════════════════
    # SKU TOP (26년 누적, 채널별)
    # ════════════════════
    sku_by_channel = {}
    for ch, grp in df26.groupby('채널_대분류'):
        top = grp.groupby('베이스품명')['원화판매금액'].sum().sort_values(ascending=False).head(5)
        sku_by_channel[str(ch)] = [
            {'name': k, 'v26': round(float(v) / 1e6, 2)} for k, v in top.items()
        ]

    # ════════════════════
    # 집계 기준일 정보
    # ════════════════════
    last_month = int(df26['마감월'].max())
    last_date = str(df26['실적일자'].max())[:10]
    months_done = sorted(df26['마감월'].unique().tolist())

    # ════════════════════
    # GPM 계산
    # ════════════════════
    gpm26 = {}
    gpm25 = {}
    for m in months_done:
        grp = df26[df26['마감월'] == m]
        sales = grp['원화판매금액'].sum()
        gp    = grp['이익액'].sum()
        gpm26[m] = round(gp / sales * 100, 2) if sales > 0 else None

    for m in sorted(df25['마감월_int'].unique()):
        grp = df25[df25['마감월_int'] == m]
        sales = grp['추정 매출액'].sum()
        cost  = grp['추정 매출원가'].sum()
        gp    = sales - cost
        gpm25[m] = round(gp / sales * 100, 2) if sales > 0 else None

    # ════════════════════
    # 사업계획 (대시보드_전사에서 읽기)
    # ════════════════════
    df_plan = pd.read_excel(xl, sheet_name='대시보드_전사', header=None)
    plan_monthly = {}
    plan_cum = {}
    try:
        plan_row = None
        for i, row in df_plan.iterrows():
            if '사업계획' in str(row.values) and '월별' in str(df_plan.iloc[i+1].values if i+1 < len(df_plan) else ''):
                plan_row = i + 1
                break
        if plan_row is None:
            for i, row in df_plan.iterrows():
                if any('사업계획' in str(v) for v in row.values):
                    for j in range(i, min(i+5, len(df_plan))):
                        r = df_plan.iloc[j]
                        if '월별' in str(r.values):
                            for col_idx in range(5, 17):
                                try:
                                    v = float(df_plan.iloc[j, col_idx])
                                    if v > 1000:
                                        plan_monthly[col_idx - 4] = round(v, 2)
                                except:
                                    pass
                            break
    except:
        pass

    # 대시보드_전사에서 사업계획 직접 추출
    for i, row in df_plan.iterrows():
        vals = [str(v) for v in row.values]
        if '사업계획' in vals or any('계획' in v for v in vals):
            for j in range(i, min(i+4, len(df_plan))):
                r = df_plan.iloc[j]
                if any('월별' in str(v) for v in r.values):
                    for col_idx in range(5, 17):
                        try:
                            v = float(r.iloc[col_idx])
                            if 5000 < v < 20000:
                                plan_monthly[col_idx - 4] = round(v, 2)
                        except:
                            pass
                    break

    # fallback: 알려진 계획값 사용
    if not plan_monthly:
        plan_monthly = {1: 9105, 2: 8166, 3: 8794, 4: 10084,
                        5: 10394, 6: 11158, 7: 10534, 8: 10544,
                        9: 11517, 10: 10694, 11: 10694, 12: 11527}

    # ════════════════════
    # 최종 JSON 조립
    # ════════════════════
    data = {
        "meta": {
            "last_date": last_date,
            "last_month": last_month,
            "months_done": months_done,
            "generated_at": pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')
        },
        "monthly": {
            "sales_26": monthly26,
            "sales_25": monthly25,
            "gp_26": monthly_gp26,
            "gp_25": monthly_gp25_real,
            "gpm_26": {int(k): v for k, v in gpm26.items()},
            "gpm_25": {int(k): v for k, v in gpm25.items()},
            "domestic_26": domestic26,
            "domestic_25": domestic25,
            "overseas_26": overseas26,
            "overseas_25": overseas25,
            "plan": plan_monthly
        },
        "brand": {
            "sales_26": brand26,
            "sales_25": brand25,
            "daiso_26": daiso26,
            "daiso_25": daiso25
        },
        "channel": {
            "acc_26": channel_acc26,
            "acc_25": channel_acc25
        },
        "sku_by_channel": sku_by_channel
    }

    # JSON 저장
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 변환 완료: {OUTPUT_PATH}")
    print(f"  기준일: {last_date}")
    print(f"  집계 월: {months_done}")
    print(f"  전사 누적 매출: {sum(monthly26.values()):.1f} 백만원")
    print(f"  전사 누적 GP: {sum(monthly_gp26.values()):.1f} 백만원")

if __name__ == '__main__':
    convert()
