# -*- coding: utf-8 -*-
"""진주 고장 대시보드 V2 - 모바일 우선 / 경량 JSON 조회."""
import json
import html
from pathlib import Path
from datetime import datetime
import streamlit as st
import pandas as pd

st.set_page_config(page_title="GN Fault Insight", page_icon="📡", layout="wide", initial_sidebar_state="collapsed")
DATA = Path(__file__).parent / 'data' / 'processed.json'

CSS = """
<style>
  .block-container{max-width:1120px;padding-top:1.3rem!important;padding-bottom:3rem!important}
  html,body,[class*='css']{font-family:system-ui,-apple-system,'Apple SD Gothic Neo','Malgun Gothic',sans-serif}
  .hero{background:linear-gradient(120deg,#10243f,#1865aa);color:white;border-radius:18px;padding:20px 24px;margin:6px 0 16px;box-shadow:0 12px 25px #0a254015}
  .hero .eyebrow{letter-spacing:.11em;text-transform:uppercase;font-size:11px;color:#b7d8f8;font-weight:700}
  .hero h1{font-size:26px!important;color:#fff!important;line-height:1.2!important;margin:6px 0 5px!important;padding:0!important}
  .hero .sub{font-size:13px;color:#e0ecfc}
  .stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:7px 0 16px}
  .stat{border:1px solid #e7edf5;background:#fff;border-radius:13px;padding:13px 15px;box-shadow:0 4px 13px #1b386208}
  .stat .label{font-size:12px;color:#69788e;font-weight:650}.stat .value{font-size:26px;font-weight:850;line-height:1.3;margin-top:7px;color:#173960}
  .stat .note{font-size:10px;color:#8b9bad;margin-top:3px}
  .section-head{font-weight:800;font-size:17px;margin:14px 0 12px}
  .alarm{background:#fff;border:1px solid #e8edf4;border-left:4px solid #9aa8b7;padding:14px;border-radius:12px;margin:0 0 10px;box-shadow:0 2px 12px #17396008}
  .alarm.red{border-left-color:#df4c5c}.alarm.orange{border-left-color:#e9a02c}.alarm.green{border-left-color:#28a574}.alarm.blue{border-left-color:#3884db}
  .alarm .head{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}
  .alarm .name{font-size:15px;font-weight:800;color:#19334c;overflow-wrap:anywhere}
  .alarm .region{font-size:12px;color:#8391a3;white-space:nowrap}
  .alarm .port{display:block;margin:6px 0;font-size:12px;color:#51667d;overflow-wrap:anywhere}
  .alarm .detail{margin:7px 0 0;font-size:12px;color:#51667d;overflow-wrap:anywhere}
  .alarm .footer{display:flex;align-items:center;justify-content:space-between;gap:6px;margin-top:10px;font-size:11px;color:#8a99a8}
  .tag{display:inline-block;border-radius:7px;padding:3px 8px;font-size:11px;font-weight:700;line-height:1.5;background:#e8eff9;color:#2864a2;margin-right:4px}
  .tag.red{background:#ffedf0;color:#b73549}.tag.orange{background:#fff3d9;color:#96600e}.tag.green{background:#e3f8ef;color:#138356}
  .tag.grey{background:#eff2f6;color:#6d7989}
  [data-testid='stDataFrame']{border-radius:12px;overflow:hidden}
  @media(prefers-color-scheme:dark){.stat,.alarm{background:#1b2b40;border-color:#354760}.alarm .name{color:#edf4fc}.alarm .port,.alarm .detail{color:#c3d1df}.stat .value{color:#dcecff}.stat .label{color:#b4c6d9}}
  @media(max-width:700px){
    .block-container{padding:.65rem .65rem 2rem!important;padding-top:1.4rem!important}
    .hero{border-radius:14px;padding:19px 16px;margin-bottom:13px}
    .hero h1{font-size:21px!important}
    .stats{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}
    .stat{padding:12px}.stat .value{font-size:25px}
    .alarm{padding:12px;margin-bottom:8px}.alarm .name{font-size:14px}
    [data-testid='stHorizontalBlock']{gap:.4rem}
  }
</style>
"""
st.markdown(CSS,unsafe_allow_html=True)

def esc(v):
    return html.escape(str(v if v is not None else ''),quote=True)

@st.cache_data(ttl=120,show_spinner=False)
def get_data(last_modified):
    with DATA.open('r',encoding='utf-8') as fp:
        return json.load(fp)

if not DATA.exists():
    st.error('데이터 처리 파일이 아직 생성되지 않았어요. GitHub Actions → Build daily alarm snapshot 실행 결과를 확인해 주세요.')
    st.info('새 저장소에 data/daily.xlsx 파일을 올리고 GitHub Actions 작업이 완료되면 자동 생성됩니다.')
    st.stop()

if st.button('↻ 최신 자료 확인',key='reload'):
    st.cache_data.clear()
    st.rerun()

db = get_data(DATA.stat().st_mtime_ns)
meta=db.get('meta',{})
st.markdown(f"""<div class="hero"><div class="eyebrow">NETWORK OPERATIONS · FIELD INSIGHT</div>
<h1>📡 GN Fault Insight</h1><div class="sub">진주 고장 현황 · 기준일 {esc(meta.get('snapshot_date','-'))} · 처리시각 {esc(meta.get('generated_at','-'))}</div></div>""",unsafe_allow_html=True)

views = ['🔔 금일 알람','📊 종합 현황','📋 전체 RAW','📡 중계기·MIBOS','🕐 13시 알람','🗂️ 기타 알람']
view=st.selectbox('조회 메뉴',views,key='menu',label_visibility='collapsed')
summary=db.get('summary',{})

def kpis(items):
    parts=[]
    for label,value,color,note in items:
        parts.append(f'<div class="stat"><div class="label">{esc(label)}</div><div class="value" style="color:{color}">{esc(value)}</div><div class="note">{esc(note)}</div></div>')
    st.markdown('<div class="stats">'+''.join(parts)+'</div>',unsafe_allow_html=True)

def show_table(items,cols,maxrows=300):
    if not items:
        st.info('표시할 데이터가 없습니다.')
        return
    frame=pd.DataFrame(items)
    frame=frame[[c for c in cols if c in frame.columns]]
    st.caption(f'총 {len(frame):,}건 · 화면에는 최대 {maxrows:,}건 표시')
    st.dataframe(frame.head(maxrows),use_container_width=True,hide_index=True,height=min(540,60+min(maxrows,len(frame))*36))
    st.download_button('⬇️ 전체 자료 CSV 저장',data=frame.to_csv(index=False).encode('utf-8-sig'),file_name='GN_Fault_Insight.csv',mime='text/csv',use_container_width=True)

if view=='🔔 금일 알람':
    kpis([
        ('전체',summary.get('today_total',0),'#2368be','금일 데이터'),
        ('신규',summary.get('today_new',0),'#dc4252','전일 미포함'),
        ('점검·조치',summary.get('today_checked',0),'#d28a11','고장구분 기재'),
        ('미점검',summary.get('today_unchecked',0),'#7956c2','전일 존재 · 미기재'),
    ])
    records=db.get('today',[])
    c1,c2=st.columns([3,1])
    search=c1.text_input('검색',placeholder='🔎 장비명 · PORT · 지역 · 알람',label_visibility='collapsed')
    sysfilter=c2.selectbox('시스템',['전체','5G','LTE'],label_visibility='collapsed')
    categories=st.radio('표시 범위',['전체','신규','미점검','점검·조치'],horizontal=True,label_visibility='collapsed')
    def ok(rec):
        hay=' '.join(str(rec.get(k,'')) for k in ('장비명','Port','시군구','알람','복구/미복구 상세내역'))
        return (not search or search.lower() in hay.lower()) and (sysfilter=='전체' or rec.get('시스템')==sysfilter) and (categories=='전체' or rec.get('현장상태')==categories)
    selected=[x for x in records if ok(x)]
    st.caption(f'검색 결과 {len(selected):,}건 · 원본 기준일 {esc(meta.get("snapshot_date",""))}')
    cards=[]
    for rec in selected[:220]:
        status=rec.get('현장상태','기타')
        color={'신규':'red','미점검':'orange','점검·조치':'blue'}.get(status,'green')
        alarm=rec.get('알람') or '확인필요'
        cards.append(f"""<div class="alarm {color}"><div class="head"><span class="name">{esc(rec.get('장비명') or rec.get('Port') or '이름 없음')}</span><span class="region">{esc(rec.get('시군구'))}</span></div>
        <span class="port">{esc(rec.get('시스템'))} · PORT: {esc(rec.get('Port'))}</span>
        <span class="tag {color}">{esc(status)}</span><span class="tag">{esc(alarm)}</span>
        <div class="detail">{esc(rec.get('복구/미복구 상세내역'))}</div>
        <div class="footer"><span>알람 근거: {esc(rec.get('알람근거') or '미확인')}</span><span>{esc(rec.get('발생시각'))}</span></div></div>""")
    st.markdown(''.join(cards) if cards else '<p>표시할 알람이 없습니다.</p>',unsafe_allow_html=True)
    if len(selected)>220: st.info('빠른 화면에서는 220건까지만 표시합니다. 전체 RAW 메뉴에서 검색·다운로드할 수 있어요.')
    df=pd.DataFrame(selected)
    if not df.empty:
        st.download_button('⬇️ 금일 알람 전체 다운로드',data=df.to_csv(index=False).encode('utf-8-sig'),file_name=f"금일알람_{meta.get('snapshot_date','')}.csv",mime='text/csv',use_container_width=True)

elif view=='📊 종합 현황':
    kpis([
        ('누적 발생',summary.get('all_total',0),'#2368be','OOS 중복 제거'),
        ('금일 관제',summary.get('today_total',0),'#dc4252','최신 스냅샷'),
        ('반복 PORT',summary.get('repeated_ports',0),'#d28a11','2회 이상'),
        ('과거 이력',summary.get('history_days',0),'#7956c2','저장 날짜 수'),
    ])
    st.markdown('<div class="section-head">최근 고장 추이</div>',unsafe_allow_html=True)
    trend=db.get('trends',[])
    if trend:
        f=pd.DataFrame(trend)
        st.bar_chart(f.set_index('날짜')['건수'],height=250)
    st.markdown('<div class="section-head">반복 PORT TOP 15</div>',unsafe_allow_html=True)
    repeat=db.get('repeated',[])
    show_table(repeat,['Port','횟수','최근발생','장비명'],maxrows=15)
    st.markdown('<div class="section-head">복구 작업 유형</div>',unsafe_allow_html=True)
    work=db.get('work_types',[])
    if work:
        st.bar_chart(pd.DataFrame(work).set_index('작업유형')['건수'],height=290,horizontal=True)

elif view=='📋 전체 RAW':
    kpis([
        ('누적 발생',summary.get('all_total',0),'#2368be','PORT+발생시각 중복 제거'),
        ('반복 PORT',summary.get('repeated_ports',0),'#d28a11','2회 이상'),
        ('이력 보존일',summary.get('history_days',0),'#277c57','날짜별 저장'),
        ('금일',summary.get('today_total',0),'#dc4252','최신 스냅샷'),
    ])
    rows=db.get('all_oos',[])
    q=st.text_input('전체 RAW 검색',placeholder='PORT · 장비명 · 복구내용 · 알람')
    if q:
        rows=[r for r in rows if q.lower() in ' '.join(str(v) for v in r.values()).lower()]
    show_table(rows,['본부','발생시각','장비명','Port','시스템','시군구','고장구분','복구/미복구 상세내역','알람','다발횟수'])

elif view=='📡 중계기·MIBOS':
    relay=db.get('sheets',{}).get('중계기 및 MIBOS 알람',[])
    grems=db.get('sheets',{}).get('gREMS',[])
    kpis([('중계기·MIBOS',len(relay),'#2368be','원본 시트'),('gREMS',len(grems),'#7956c2','원본 시트')])
    st.markdown('#### 중계기 및 MIBOS 알람')
    show_table(relay,['시스템 구분','경과일','중계기명','알람등급','알람내역','기지국명'])
    st.markdown('#### gREMS')
    show_table(grems,['알람등급','알람발생일자','중계기명','알람명','서비스 그룹명'])

elif view=='🕐 13시 알람':
    df=db.get('sheets',{}).get('13시 알람 공유',[])
    kpis([('13시 공유 알람',len(df),'#2368be','원본 시트')])
    show_table(df,['장비명','알람','발생시각','시군구','발생위치'])

else:
    for key in ('RMS_A망 미복구','RMS_DACS 미복구','RMS_통합RCU미복구'):
        items=db.get('sheets',{}).get(key,[])
        st.markdown(f'#### {key} · {len(items):,}건')
        show_table(items,list(items[0])[:12] if items else [])

st.caption('데이터는 매일 Excel 업로드 후 GitHub Actions에서 자동 생성됩니다. 서버 처리 실패 시 이전 자료가 표시될 수 있으니 기준일을 확인하세요.')
