"""엑셀 → JSON 변환 스크립트 (원본 대시보드 완전 연동 버전)"""
import pandas as pd
import json
import sys
from pathlib import Path

EXCEL_PATH = Path("data/실적데이터.xlsx")
OUTPUT_PATH = Path("dashboard_data.json")

def safe_round(v, d=2):
    try:
        if pd.isna(v): return None
        return round(float(v), d)
    except:
        return None

def load_db26(xl):
    df = pd.read_excel(xl, sheet_name='매출 DB_2026', header=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df[df['원화판매금액'].notna()]

def load_db25(xl):
    df = pd.read_excel(xl, sheet_name='매출 DB_2025', header=4)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df['추정 매출액'].notna()]
    df['마감월_int'] = df['마감월'].apply(
        lambda x: int(str(x).replace('25-','').strip()) if pd.notna(x) and str(x).startswith('25-') else None)
    return df[df['마감월_int'].notna()]

def channel_norm(ch):
    if not ch or str(ch)=='nan': return '기타'
    ch=str(ch)
    for p in ['CIS권','동남아권','중국권','일본권','북미권','동유럽권','Global권','중동권','기타유럽권']:
        if ch.startswith(p): return p
    return ch

def brand_26(line):
    line=str(line)
    HERO=['펩타이드 9','피디알엔','레티놀 콜라겐','피토 이엑스 피디알엔','영시카 피디알엔','레티날']
    CHAMP=['레드 락토 콜라겐','멜라논 엑스','프리미엄 콜라겐','그린 시카 콜라겐','엑스트라 슈퍼 9 플러스','EGF','비건 비타민','시카놀 B5']
    if any(h in line for h in HERO): return 'hero'
    if any(c in line for c in CHAMP): return 'champion'
    return 'other'

def brand_25(row):
    line=str(row.get('라인명',''))
    HERO=['펩타이드 9','피디알엔','레티날','레티놀','피토 이엑스 피디알엔']
    CHAMP=['레드 락토 콜라겐','멜라논','프리미엄 콜라겐','그린 시카 콜라겐','엑스트라 슈퍼 9','EGF','비건 비타민','시카놀']
    if any(h in line for h in HERO): return 'hero'
    if any(c in line for c in CHAMP): return 'champion'
    return 'other'

def convert():
    if not EXCEL_PATH.exists():
        print(f"ERROR: {EXCEL_PATH} 파일이 없습니다."); sys.exit(1)

    xl = pd.ExcelFile(EXCEL_PATH)
    print(f"파일 로드 완료: {EXCEL_PATH}")

    df26 = load_db26(xl)
    df25 = load_db25(xl)
    df26['채널'] = df26['채널그룹'].apply(channel_norm)
    df26['브랜드'] = df26['라인'].apply(brand_26)
    df26['다이소'] = df26['유통구조'].apply(lambda x: '다이소' in str(x))
    df25['채널'] = df25['유통구조'].apply(channel_norm)
    df25['브랜드'] = df25.apply(brand_25, axis=1)
    df25['다이소'] = df25['구분2'].apply(lambda x: '다이소' in str(x))

    months26 = sorted(df26['마감월'].dropna().unique().astype(int).tolist())
    last_month = max(months26)
    last_date = str(df26['실적일자'].max())[:10]

    try:
        from datetime import datetime; import calendar
        d = datetime.strptime(last_date,'%Y-%m-%d')
        days_in_month = calendar.monthrange(d.year,d.month)[1]
        days_elapsed = d.day; ratio = days_elapsed/days_in_month
    except:
        days_elapsed=3; days_in_month=30; ratio=0.1

    # 월별 집계
    s26={int(m):safe_round(v/1e6) for m,v in df26.groupby('마감월')['원화판매금액'].sum().items()}
    s25={int(m):safe_round(v/1e6) for m,v in df25.groupby('마감월_int')['추정 매출액'].sum().items()}
    gp26={int(m):safe_round(v/1e6) for m,v in df26.groupby('마감월')['이익액'].sum().items()}
    gp25={}; gpm25={}
    for m,grp in df25.groupby('마감월_int'):
        sales=grp['추정 매출액'].sum(); cost=grp['추정 매출원가'].sum()
        gp25[int(m)]=safe_round((sales-cost)/1e6)
        gpm25[int(m)]=safe_round((sales-cost)/sales*100) if sales>0 else None
    gpm26={}
    for m,grp in df26.groupby('마감월'):
        sales=grp['원화판매금액'].sum(); gp=grp['이익액'].sum()
        gpm26[int(m)]=safe_round(gp/sales*100) if sales>0 else None

    dom26={int(m):safe_round(v/1e6) for m,v in df26[df26['국내/해외']=='국내'].groupby('마감월')['원화판매금액'].sum().items()}
    ov26={int(m):safe_round(v/1e6) for m,v in df26[df26['국내/해외']=='해외'].groupby('마감월')['원화판매금액'].sum().items()}
    dom25={}; ov25={}
    for m,grp in df25.groupby('마감월_int'):
        dom25[int(m)]=safe_round(grp[grp['지역']=='국내']['추정 매출액'].sum()/1e6)
        ov25[int(m)]=safe_round(grp[grp['지역']!='국내']['추정 매출액'].sum()/1e6)

    b26={}
    for (m,b),grp in df26.groupby(['마감월','브랜드']): b26.setdefault(b,{})[int(m)]=safe_round(grp['원화판매금액'].sum()/1e6)
    b25={}
    for (m,b),grp in df25.groupby(['마감월_int','브랜드']): b25.setdefault(b,{})[int(m)]=safe_round(grp['추정 매출액'].sum()/1e6)
    d26b={int(m):safe_round(v/1e6) for m,v in df26[df26['다이소']].groupby('마감월')['원화판매금액'].sum().items()}
    d25b={int(m):safe_round(v/1e6) for m,v in df25[df25['다이소']].groupby('마감월_int')['추정 매출액'].sum().items()}

    ch26={str(k):safe_round(v/1e6) for k,v in df26.groupby('채널')['원화판매금액'].sum().sort_values(ascending=False).items()}
    ch25={str(k):safe_round(v/1e6) for k,v in df25.groupby('채널')['추정 매출액'].sum().sort_values(ascending=False).items()}
    ch_m26={}
    for (m,ch),grp in df26.groupby(['마감월','채널']): ch_m26.setdefault(str(ch),{})[int(m)]=safe_round(grp['원화판매금액'].sum()/1e6)

    sku_by_ch={}
    for ch,grp in df26.groupby('채널'):
        top=grp.groupby('베이스품명')['원화판매금액'].sum().sort_values(ascending=False).head(5)
        sku_by_ch[str(ch)]=[{'name':k,'v26':safe_round(v/1e6)} for k,v in top.items()]

    plan={1:9105,2:8166,3:8794,4:10084,5:10394,6:11158,7:10534,8:10544,9:11517,10:10694,11:10694,12:11527}
    try:
        df_plan=pd.read_excel(xl,sheet_name='대시보드_전사',header=None)
        for i,row in df_plan.iterrows():
            if any('사업계획' in str(v) for v in row.values):
                for j in range(i,min(i+5,len(df_plan))):
                    r=df_plan.iloc[j]
                    if any('월별' in str(v) for v in r.values):
                        tmp={col_idx-4:round(float(r.iloc[col_idx]),2) for col_idx in range(5,17) if 5000<float(r.iloc[col_idx] if pd.notna(r.iloc[col_idx]) else 0)<20000}
                        if len(tmp)>=3: plan=tmp
                        break
    except: pass

    # KPI 계산
    acc_s26=sum(s26.get(m,0) for m in months26)
    acc_s25=sum(s25.get(m,0) for m in months26)
    acc_gp26=sum(gp26.get(m,0) for m in months26)
    acc_gp25=sum(gp25.get(m,0) for m in months26)
    acc_plan=sum(plan.get(m,0) for m in months26)
    cur=s26.get(last_month,0); cur25=s25.get(last_month,0)
    cur_gp=gp26.get(last_month,0); cur_gp25=gp25.get(last_month,0)
    cur_gpm=gpm26.get(last_month,0); cur_gpm25=gpm25.get(last_month,0)
    proj=safe_round(cur/ratio) if ratio>0 else cur
    proj_gp=safe_round(cur_gp/ratio) if ratio>0 else cur_gp
    acc_gpm26=safe_round(acc_gp26/acc_s26*100) if acc_s26>0 else None
    acc_gpm25=safe_round(acc_gp25/acc_s25*100) if acc_s25>0 else None

    # 브랜드별 누적
    hero26=sum(b26.get('hero',{}).get(m,0) for m in months26)
    hero25=sum(b25.get('hero',{}).get(m,0) for m in months26)
    champ26=sum(b26.get('champion',{}).get(m,0) for m in months26)
    champ25=sum(b25.get('champion',{}).get(m,0) for m in months26)
    daiso26=sum(d26b.get(m,0) for m in months26)
    daiso25=sum(d25b.get(m,0) for m in months26)
    other26=sum(b26.get('other',{}).get(m,0) for m in months26)
    other25=sum(b25.get('other',{}).get(m,0) for m in months26)

    data={
        "meta":{"last_date":last_date,"last_month":last_month,"months_done":months26,
                "days_elapsed":days_elapsed,"days_in_month":days_in_month,
                "ratio":safe_round(ratio,4),"generated_at":pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')},
        "kpi":{
            "acc_sales_26":safe_round(acc_s26),"acc_sales_25":safe_round(acc_s25),
            "cur_month_sales":safe_round(cur),"cur_month_sales_25":safe_round(cur25),
            "proj_month_sales":safe_round(proj),
            "acc_plan":safe_round(acc_plan),"cur_plan":plan.get(last_month,0),
            "yoy_acc":safe_round((acc_s26-acc_s25)/acc_s25*100) if acc_s25>0 else None,
            "yoy_month":safe_round((cur-cur25)/cur25*100) if cur25>0 else None,
            "achv_acc":safe_round(acc_s26/acc_plan*100) if acc_plan>0 else None,
            "achv_month":safe_round(cur/plan.get(last_month,1)*100),
            "acc_gp_26":safe_round(acc_gp26),"acc_gp_25":safe_round(acc_gp25),
            "acc_gpm_26":acc_gpm26,"acc_gpm_25":acc_gpm25,
            "cur_gp_26":safe_round(cur_gp),"cur_gp_25":safe_round(cur_gp25),
            "cur_gpm_26":safe_round(cur_gpm),"cur_gpm_25":safe_round(cur_gpm25),
            "proj_gp":safe_round(proj_gp),
            "hero_acc_26":safe_round(hero26),"hero_acc_25":safe_round(hero25),
            "champ_acc_26":safe_round(champ26),"champ_acc_25":safe_round(champ25),
            "daiso_acc_26":safe_round(daiso26),"daiso_acc_25":safe_round(daiso25),
            "other_acc_26":safe_round(other26),"other_acc_25":safe_round(other25),
        },
        "monthly":{"sales_26":s26,"sales_25":s25,"gp_26":gp26,"gp_25":gp25,
                   "gpm_26":gpm26,"gpm_25":gpm25,"domestic_26":dom26,"domestic_25":dom25,
                   "overseas_26":ov26,"overseas_25":ov25,"plan":plan},
        "brand":{"sales_26":b26,"sales_25":b25,"daiso_26":d26b,"daiso_25":d25b},
        "channel":{"acc_26":ch26,"acc_25":ch25,"monthly_26":ch_m26},
        "sku_by_channel":sku_by_ch
    }

    with open(OUTPUT_PATH,'w',encoding='utf-8') as f:
        json.dump(data,f,ensure_ascii=False,indent=2)

    print(f"\n✅ 변환 완료: {OUTPUT_PATH}")
    print(f"  기준일: {last_date} ({days_elapsed}/{days_in_month}일, {ratio*100:.0f}%)")
    print(f"  누적 매출: {acc_s26:.1f} / GP: {acc_gp26:.1f} 백만원")
    print(f"  {last_month}월: {cur:.1f} → 예상 {proj:.1f} 백만원")

if __name__=='__main__':
    convert()
