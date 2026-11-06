"""Jinju Fault Insight V3. Read-only UI: no Excel parsing, HTTP requests or history writes."""
from __future__ import annotations
import gzip
import html
import io
import json
import math
from pathlib import Path
from datetime import datetime, date
from zoneinfo import ZoneInfo
import pandas as pd
import streamlit as st

ROOT=Path(__file__).resolve().parent
DATA=ROOT/'data/v3'
st.set_page_config(page_title='JINJU · 고장 현황',page_icon='📡',layout='wide',initial_sidebar_state='collapsed')
CSS='<style>\nhtml,body,[class*="css"]{font-family:system-ui,-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif}\n.block-container{max-width:1120px;padding:2rem 1.3rem 3rem!important}\n[data-testid="stSidebar"],[data-testid="stSidebarCollapsedControl"]{display:none!important}\n.j-apphead{font-size:20px;font-weight:750;letter-spacing:-.5px;line-height:1.4;padding:4px 0 12px;color:inherit}.j-apphead small{display:block;font-size:12px;font-weight:450;opacity:.65;letter-spacing:0;margin-bottom:3px}\n.j-title{font-size:22px!important;line-height:1.35!important;margin:8px 0 4px!important;padding:0!important;color:inherit}.j-subtitle{font-size:12px;opacity:.65;margin-bottom:15px}\n.j-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:0;background:rgba(128,128,128,.045);border:1px solid rgba(128,128,128,.17);border-radius:12px;padding:14px 6px;margin:8px 0 12px;text-align:center}.j-kpi{min-width:0;padding:0 5px}.j-kpi .label{font-size:12px;opacity:.65;padding-bottom:8px;border-bottom:1px solid rgba(128,128,128,.15);white-space:nowrap}.j-kpi .value{font-size:28px;font-weight:750;line-height:1.3;margin-top:9px;color:#3984d5}.j-kpi:nth-child(2) .value{color:#e95158}.j-kpi:nth-child(3) .value{color:#d98a2d}.j-kpi:nth-child(4) .value{color:#a36cde}.j-kpi .hint{display:none}\n.j-cards{display:flex;flex-direction:column;gap:9px;margin:5px 0 12px}.j-alarm{background:rgba(128,128,128,.055);border:1px solid rgba(128,128,128,.12);border-left:4px solid #e95158;border-radius:10px;padding:13px 14px;line-height:1.5;color:inherit;min-width:0}.j-alarm.bp{border-left-color:#3984d5}.j-alarm.done{border-left-color:#39a77d}.j-alarm.check{border-left-color:#d98a2d}.j-alarm.other{border-left-color:#a36cde}\n.j-cardhead{display:flex;justify-content:space-between;align-items:flex-start;gap:10px}.j-device{font-size:14px;font-weight:700;overflow-wrap:anywhere;min-width:0}.j-region{font-size:12px;opacity:.6;white-space:nowrap;flex-shrink:0;padding-top:2px}.j-sys{display:inline-block;padding:2px 8px;border-radius:14px;background:rgba(128,128,128,.12);font-size:11px;font-weight:700;color:#3984d5;margin-right:5px;vertical-align:middle}.j-sys.lte{color:#d98a2d}\n.j-portrow{margin-top:6px;min-width:0;overflow-wrap:anywhere}.j-port{font-family:ui-monospace,SFMono-Regular,Consolas,monospace;font-size:12px;border-radius:5px;background:rgba(128,128,128,.12);padding:2px 6px;overflow-wrap:anywhere;white-space:normal}\n.j-alarmline{margin-top:8px}.j-pill{display:inline-block;font-size:12px;font-weight:650;border-radius:7px;padding:3px 8px;background:rgba(128,128,128,.12);color:inherit;overflow-wrap:anywhere;max-width:100%}.j-pill.link{color:#e95158}.j-pill.power{color:#d98a2d}.j-pill.fru{color:#a36cde}.j-pill.sw{color:#3984d5}.j-pill.status{color:#3984d5;font-size:11px}.j-pill.status.warn{color:#d98a2d}.j-pill.status.done{color:#39a77d}.j-pill.status.new{color:#e95158}\n.j-memo{font-size:12px;opacity:.75;margin-top:7px;overflow-wrap:anywhere}.j-cardfooter{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-top:9px;flex-wrap:wrap}.j-time{font-size:11px;opacity:.55;margin-left:auto;white-space:nowrap}\n.j-heading{font-size:16px;font-weight:700;margin:10px 0 12px}.j-heading small{font-size:12px;font-weight:400;opacity:.6;margin-left:10px}\n[data-testid="stDataFrame"]{border-radius:9px;overflow:hidden}\n.stButton button,.stDownloadButton button{border-radius:9px;font-size:13px;min-height:40px}\n.stDownloadButton button{background:#286cb9;color:white;border-color:#286cb9;width:100%}\n[data-testid="stCaptionContainer"] p{font-size:12px}\n[data-testid="stPills"]{margin:4px 0 10px}\n.st-key-today_filters [data-testid="stHorizontalBlock"]{flex-wrap:nowrap;gap:10px}.st-key-today_filters [data-testid="stColumn"]{min-width:0!important;flex:1 1 0!important}\n@media(max-width:700px){.block-container{padding:1.4rem .65rem 2rem!important}.j-apphead{font-size:19px;padding-top:2px}.j-title{font-size:20px!important}.j-kpis{padding:12px 4px}.j-kpi{padding:0 3px}.j-kpi .label{font-size:11px}.j-kpi .value{font-size:25px}.j-alarm{padding:12px 11px}.j-device{font-size:14px}.j-region{font-size:11px}.j-port{font-size:12px}.j-time{font-size:11px}.j-cards{gap:8px}}\n</style>'
st.markdown(CSS,unsafe_allow_html=True)

def esc(v):return html.escape(str(v if v is not None else ''),quote=True)

@st.cache_data(max_entries=6,show_spinner=False)
def read_data(path,mtime,size):
    path=Path(path)
    content=path.read_bytes()
    if path.suffix=='.gz':content=gzip.decompress(content)
    return json.loads(content)

def read(path):
    stat=path.stat();return read_data(str(path),stat.st_mtime_ns,stat.st_size)

def kpis(items):
    cards=[]
    for label,value,hint,focus in items:
        val=f'{value:,}' if isinstance(value,(int,float)) else str(value)
        cards.append(f'<div class="j-kpi {"focus" if focus else ""}"><div class="label">{esc(label)}</div><div class="value">{esc(val)}</div><div class="hint">{esc(hint)}</div></div>')
    st.markdown('<div class="j-kpis">'+''.join(cards)+'</div>',unsafe_allow_html=True)

def heading(title,note=''):
    st.markdown(f'<div class="j-heading">{esc(title)}<small>{esc(note)}</small></div>',unsafe_allow_html=True)

def chart(frame,label,value,horizontal=False):
    if frame.empty:st.info('집계할 자료가 없습니다.');return
    import plotly.express as px
    fig=px.bar(frame,x=value if horizontal else label,y=label if horizontal else value,orientation='h' if horizontal else 'v',text=value,color_discrete_sequence=['#087e83'])
    fig.update_layout(height=max(240,min(440,45*len(frame))) if horizontal else 240,paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',margin=dict(l=6,r=25,t=10,b=10),font=dict(family='system-ui, sans-serif',size=11),showlegend=False,xaxis_title='',yaxis_title='')
    fig.update_traces(textposition='outside',cliponaxis=False)
    fig.update_xaxes(showgrid=horizontal,gridcolor='#edf1f4');fig.update_yaxes(showgrid=not horizontal,gridcolor='#edf1f4',categoryorder='total ascending')
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})

def export(frame,key,name,context=()):
    # Large CSV conversion only on explicit request; never on normal filtering/rerun.
    revision=tuple((p.stat().st_mtime_ns,p.stat().st_size) for p in [DATA/'current.json',DATA/'events.json.gz',DATA/'raw_events.json.gz'] if p.exists())
    signature=(revision,repr(context),hash(tuple(frame.index)),frame.shape,tuple(frame.columns))
    if st.button('CSV 다운로드 준비',key='prepare_'+key):
        st.session_state['export_'+key]=(signature,frame.to_csv(index=False).encode('utf-8-sig'))
    ready=st.session_state.get('export_'+key)
    if ready and ready[0]==signature:
        st.download_button('CSV 저장',ready[1],file_name=name,mime='text/csv',key='download_'+key)

def paged_table(frame,key,cols=None):
    if frame.empty:st.info('조회 결과가 없습니다.');return
    if cols:frame=frame[[c for c in cols if c in frame.columns]]
    pages=max(1,math.ceil(len(frame)/100))
    page=st.number_input('페이지 · 100건씩 표시',min_value=1,max_value=pages,value=1,key=key+'_'+str(pages)) if pages>1 else 1
    st.caption(f'전체 {len(frame):,}건 · {int(page):,}/{pages:,}페이지')
    # Only the visible page is sent to the mobile browser.
    display=frame.iloc[(int(page)-1)*100:int(page)*100].copy()
    for col in display:
        if display[col].dtype=='object':display[col]=display[col].map(lambda v:', '.join(map(str,v)) if isinstance(v,list) else str(v) if v is not None else '')
    st.dataframe(display,use_container_width=True,hide_index=True,height=min(540,75+len(display)*35))

def alarm_cards(rows):
    if not rows:st.info('표시할 알람이 없습니다.');return
    pages=max(1,math.ceil(len(rows)/100))
    page=int(st.number_input('알람 목록 페이지 · 100건씩 표시',min_value=1,max_value=pages,value=1,key='alarm_cards_page_'+str(pages))) if pages>1 else 1
    start=(page-1)*100
    visible=rows[start:start+100]
    if pages>1:st.caption(f'전체 {len(rows):,}건 중 {start+1:,}~{start+len(visible):,}건 · {page}/{pages}페이지')
    parts=[]
    for row in visible:
        state=row.get('현장상태','')
        gj=str(row.get('고장구분') or '').strip()
        cardcls='bp' if gj.upper() in ('BP','UNIT') else 'done' if gj=='복구' else 'check' if gj=='점검중' or state=='미점검' else 'other' if gj else ''
        status=gj or state
        statuscls='done' if gj=='복구' else 'warn' if state=='미점검' or gj=='점검중' else 'new' if state=='신규' else ''
        alarm=row.get('알람') or '확인필요'
        alarmcls={'Input Power Failure':'power','Link Failure':'link','No Connection':'link','FRU General Problem':'fru','SW Error':'sw'}.get(alarm,'')
        sys=row.get('시스템','')
        memo=row.get('복구/미복구 상세내역') or ''
        details=f'<div class="j-memo">📝 {esc(memo)}</div>' if memo else ''
        parts.append(f'<article class="j-alarm {cardcls}"><div class="j-cardhead"><div class="j-device"><span class="j-sys {"lte" if sys=="LTE" else ""}">{esc(sys)}</span>{esc(row.get("장비명") or row.get("Port"))}</div><span class="j-region">{esc(row.get("시군구"))}</span></div><div class="j-portrow"><span class="j-port">{esc(row.get("Port"))}</span></div><div class="j-alarmline"><span class="j-pill {alarmcls}">{esc(alarm)}</span></div>{details}<div class="j-cardfooter"><span class="j-pill status {statuscls}">{esc(status)}</span><span class="j-time">{esc(row.get("발생시각"))}</span></div></article>')
    st.markdown('<div class="j-cards">'+''.join(parts)+'</div>',unsafe_allow_html=True)

@st.cache_data(max_entries=4,show_spinner=False)
def today_download(rows):
    return pd.DataFrame(rows).drop(columns=['event_id'],errors='ignore').to_csv(index=False).encode('utf-8-sig')

def history(kind='oos'):
    path=DATA/('events.json.gz' if kind=='oos' else 'raw_events.json.gz')
    if not path.exists():st.warning('누적 자료가 아직 없습니다. 자료 처리 작업을 먼저 실행해 주세요.');return pd.DataFrame()
    data=read(path)
    return pd.DataFrame(list(data['events'].values()))

st.markdown('<div class="j-apphead"><small>📡 진주품질개선팀</small>고장 현황 대시보드</div>',unsafe_allow_html=True)
if st.button('↻ 최신 자료 확인',use_container_width=True):
    read_data.clear();today_download.clear();st.rerun()
menus={'금일 알람':'금일 알람','중계기·MIBOS':'중계기 · MIBOS','기타 알람':'기타 알람','OOS 이력':'OOS 이력','보고서':'보고서 · 알람 분석','저장 상태':'원본 · 저장 상태'}
selection=st.pills('조회 메뉴',list(menus),default='금일 알람',selection_mode='single',key='top_menu',label_visibility='collapsed')
menu=menus.get(selection,'금일 알람')

if not (DATA/'current.json').exists():
    st.title('새 대시보드 자료를 준비하고 있습니다')
    st.info('GitHub Actions에서 Preserve fault history V3가 성공하면 화면이 표시됩니다. 기존 엑셀은 같은 이름으로 계속 교체하면 됩니다.')
    st.stop()
try:db=read(DATA/'current.json')
except Exception as exc:st.error(f'자료 읽기 실패: {exc}');st.stop()
meta=db['meta'];summary=db['summary'];today=sorted(db['today'],key=lambda r:r.get('발생시각',''),reverse=True);snapshot=meta['snapshot_date']
st.caption(f'📅 기준일: {snapshot}')
titles={'금일 알람':('금일 알람','장비와 알람을 빠르게 찾고 조치 상태를 확인하세요.'),'중계기 · MIBOS':('중계기 · MIBOS','LEGACY · MIBOS · gREMS 통합 조회'),'기타 알람':('기타 알람','RMS A망 · DACS · 통합RCU 미복구 현황'),'OOS 이력':('OOS 이력','고장 발생과 날짜별 관측을 구분해 기록합니다.'),'보고서 · 알람 분석':('보고서 · 알람 분석','기간별 대표 알람과 반복 고장을 확인하세요.'),'원본 · 저장 상태':('원본 · 저장 상태','일별 원본과 같은 날 재업로드된 자료까지 보관합니다.')}
title,subtitle=titles[menu]
if menu!='금일 알람':st.markdown(f'<h1 class="j-title">{esc(title)}</h1><div class="j-subtitle">{esc(subtitle)}</div>',unsafe_allow_html=True)
if menu=='금일 알람':
    kpis([('전체',summary['today_total'],'최신 기준일 OOS',False),('신규',summary['today_new'],'직전 기준일에 없던 PORT',False),('점검·조치',summary['today_checked'],'고장구분에 내용이 있는 항목',False),('미점검',summary['today_unchecked'],'전일 존재 · 고장구분 미기재',True)])
    st.download_button(f'⬇ 금일 알람 전체 다운로드 · {len(today)}건',data=today_download(today),file_name=f'금일알람_{snapshot}.csv',mime='text/csv',use_container_width=True,key='today_csv_direct')

if menu=='금일 알람':
    q=st.text_input('장비 · PORT · 알람 검색',placeholder='🔎 장비명 · PORT · 시군구 · 알람',label_visibility='collapsed')
    with st.container(key='today_filters'):
        a,b=st.columns(2)
        system=a.selectbox('시스템',['전체','5G','LTE'],format_func=lambda v:'전체 시스템' if v=='전체' else v,label_visibility='collapsed')
        state=b.selectbox('현장 상태',['전체','신규','미점검','점검·조치','비교불가'],format_func=lambda v:'전체 상태' if v=='전체' else v,key='today_state',label_visibility='collapsed')
    rows=[r for r in today if (system=='전체' or r['시스템']==system) and (state=='전체' or r['현장상태']==state) and (not q or q.lower() in ' '.join(str(v) for v in r.values()).lower())]
    st.caption(f'표시 {len(rows):,}건 · 최근 발생순')
    alarm_cards(rows)
    frame=pd.DataFrame(rows)
    if not frame.empty:
        with st.expander('발생시각 · PORT · 조치내용 상세'):
            paged_table(frame,'today_detail',['발생시각','시스템','장비명','Port','시군구','알람','알람근거','고장구분','복구/미복구 상세내역'])
        with st.expander('집계 기준 안내'):
            st.caption('신규는 직전 기준일에 없던 PORT입니다. 점검·조치 기재와 중복될 수 있으며, 실제 복구 완료를 뜻하지 않습니다. 상세 표에서 매칭 근거를 확인할 수 있습니다.')

elif menu in ['중계기 · MIBOS','기타 알람']:
    names=['중계기 및 MIBOS 알람','gREMS'] if menu=='중계기 · MIBOS' else ['RMS_A망 미복구','RMS_DACS 미복구','RMS_통합RCU미복구']
    # One table at a time on mobile; all source columns remain available.
    names_labels={n:n.replace('_',' ') for n in names}
    chosen=st.selectbox('자료 구분',names,format_func=lambda n:names_labels[n])
    frame=pd.DataFrame(db['sheets'].get(chosen,[]))
    q=st.text_input('장비 · 알람 검색',key='aux_search')
    if q and not frame.empty:frame=frame[frame.astype(str).apply(lambda r:r.str.contains(q,case=False,regex=False).any(),axis=1)]
    days_col='경과일' if '경과일' in frame else None
    dates=next((c for c in ['알람발생일자','발생일시','발생일'] if c in frame),None)
    if not frame.empty and (days_col or dates):
        age=pd.to_numeric(frame[days_col],errors='coerce') if days_col else (pd.Timestamp(snapshot)-pd.to_datetime(frame[dates],errors='coerce')).dt.days
        frame['장기 미복구']=age.map(lambda v:'15일 이상' if pd.notna(v) and v>=15 else '')
        with st.expander(f'15일 이상 장기 미복구 · {int((age>=15).sum())}건'):
            paged_table(frame[age>=15],'overdue_'+chosen)
    paged_table(frame,'aux_'+chosen)
    if not frame.empty:export(frame,'aux_'+chosen,chosen+'.csv',(chosen,q))

elif menu=='OOS 이력':
    frame=history()
    if not frame.empty:
        kpis([('고유 발생',len(frame),'시스템 + PORT + 발생시각',False),('반복 PORT',summary['repeated_ports'],'서로 다른 발생이 2회 이상',False),('보관 원본',summary['snapshot_total'],'같은 날 수정본도 별도 보관',False),('금일 관측',summary['today_total'],'최신 OOS 기준',True)])
        a,b,c=st.columns([3,1,1]);q=a.text_input('장비 · PORT · 알람 검색',key='history_q');system=b.selectbox('시스템',['전체','5G','LTE'],key='history_system');status=c.selectbox('복구 상태',['전체','미복구 관측','복구','금일 미관측 · 복구 추정'])
        st.caption('금일 목록에 없으면 복구 추정으로 구분합니다. 실제 복구시각이 있는 경우에만 복구로 표시합니다.')
        if q:frame=frame[frame.astype(str).apply(lambda r:r.str.contains(q,case=False,regex=False).any(),axis=1)]
        if system!='전체':frame=frame[frame['시스템']==system]
        if status!='전체':frame=frame[frame['상태']==status]
        frame=frame.sort_values('발생시각',ascending=False)
        columns=['발생시각','시스템','장비명','Port','시군구','최초확인알람','최초알람근거','알람','알람근거','상태','고장구분','고장구분(중분류)','복구/미복구 상세내역','최초기준일','최종기준일','관측일수']
        paged_table(frame,'history',columns)
        export(frame[[c for c in columns if c in frame]],'history',f'OOS_누적이력_{snapshot}.csv',(q,system,status))
        if not frame.empty:
            with st.expander('개별 고장 관측 이력'):
                subset=frame.head(300).set_index('event_id',drop=False)
                eid=st.selectbox('고장 선택 · 현재 검색 결과 상위 300건',list(subset.index),format_func=lambda i:f'{subset.loc[i,"발생시각"]} · {subset.loc[i,"Port"]}')
                event=subset.loc[eid];st.write('관측 기준일: '+', '.join(event['관측일']));st.write('확인된 알람 종류: '+', '.join(event['알람종류']))
                st.caption('동일 날짜의 업로드별 세부 기록은 원본 · 저장 상태 메뉴에서 조회합니다.')

elif menu=='보고서 · 알람 분석':
    scope=st.radio('분석 자료',['OOS 고유 발생','RAW 고유 기록'],horizontal=True)
    raw=scope=='RAW 고유 기록';frame=history('raw' if raw else 'oos')
    if not frame.empty:
        time_col='eventTime' if raw else '발생시각';alarm_col='specificProblem' if raw else '최초확인알람'
        a,b,c=st.columns([1,1,1]);earliest=pd.to_datetime(frame[time_col],errors='coerce').min().date();latest=pd.to_datetime(frame[time_col],errors='coerce').max().date()
        begin=a.date_input('발생 시작일',earliest,key='begin_'+str(raw));end=b.date_input('발생 종료일',latest,key='end_'+str(raw));system=c.selectbox('시스템',['전체']+sorted(frame['시스템'].dropna().unique().tolist()),key='report_sys_'+str(raw))
        if begin>end:st.error('시작일이 종료일보다 늦습니다.');st.stop()
        chosen=frame[(frame[time_col].str[:10]>=begin.isoformat())&(frame[time_col].str[:10]<=end.isoformat())]
        if system!='전체':chosen=chosen[chosen['시스템']==system]
        if raw:
            st.caption('RAW 기록은 시스템·DU·발생시각·알람명·객체·eventType이 같으면 1건입니다. 일별 중복 관측은 제외하며, 발생/해제 메시지를 포함할 수 있어 OOS 발생 건수와 다릅니다.')
        else:
            include_legacy=st.checkbox('구 이력에서 가져온 미검증 알람도 포함',value=False)
            # Provenance filtering affects alarm rankings, never the underlying occurrence totals.
            st.caption('알람 순위는 고장별 최초 확인 대표 알람 기준입니다. 현재 RAW를 과거 고장에 소급하지 않으며, 미확인 알람은 순위에서 제외합니다.')
        ranking_base=chosen if raw or include_legacy else chosen[~chosen['최초알람근거'].str.startswith('구이력')]
        matched=ranking_base[ranking_base[alarm_col].fillna('').ne('')]
        kpis([('선택 기간 기록',len(chosen),'중복 관측 제외',False),('알람 확인',len(matched),'알람명 확인된 기록',False),('알람 종류',matched[alarm_col].nunique(),'대표 알람 기준',False),('미확인 · 미검증',len(chosen)-len(matched),'순위에서 제외',True)])
        a,b=st.columns([1.2,1])
        ranking=matched[alarm_col].value_counts().rename_axis('알람').reset_index(name='건수')
        with a:
            with st.container(border=True):heading('많이 발생한 알람 TOP 10');chart(ranking.head(10),'알람','건수',True)
        with b:
            with st.container(border=True):
                heading('월별 기록 추이')
                monthly=chosen[time_col].str[:7].value_counts().sort_index().rename_axis('월').reset_index(name='건수')
                chart(monthly,'월','건수')
        paged_table(ranking,'alarm_ranking');export(ranking,'ranking_'+str(raw),f'알람순위_{begin}_{end}.csv',(scope,begin,end,system,include_legacy if not raw else False))
        if not raw:
            heading('반복 발생 PORT')
            repeat=chosen.groupby(['시스템','Port'],dropna=False).size().reset_index(name='발생횟수').sort_values('발생횟수',ascending=False)
            paged_table(repeat[repeat['발생횟수']>1],'repeated');export(chosen.drop(columns=['관측일'],errors='ignore'),'report_oos',f'OOS_보고서기초_{begin}_{end}.csv',(begin,end,system,include_legacy))
            heading('작업 유형 · 장비 중분류')
            from fault_core import classify_work
            c1,c2=st.columns(2)
            with c1:
                works=chosen['복구/미복구 상세내역'].map(classify_work);works=works[works.ne('')].value_counts().rename_axis('작업유형').reset_index(name='건수');chart(works,'작업유형','건수',True)
            with c2:
                mid=chosen['고장구분(중분류)'];mid=mid[mid.ne('')].value_counts().rename_axis('중분류').reset_index(name='건수');chart(mid,'중분류','건수',True)
        else:export(chosen.drop(columns=['관측일'],errors='ignore'),'raw_records',f'RAW_고유기록_{begin}_{end}.csv',(begin,end,system))

elif menu=='원본 · 저장 상태':
    manifest=read(DATA/'manifest.json') if (DATA/'manifest.json').exists() else {'snapshots':[]}
    snapshots=manifest['snapshots']
    kpis([('보관 원본',len(snapshots),'파일 내용이 다른 버전',False),('보관 기준일',len({r['snapshot_date'] for r in snapshots}),'같은 날 수정본 제외',False),('OOS 누적',summary['event_total'],'고유 발생',False),('RAW 누적',summary.get('raw_event_total',0),'고유 기록',True)])
    st.write(f'최근 처리시각: {meta["generated_at"]}')
    st.caption('엑셀 업로드 후 Actions 성공까지 이전 자료가 표시됩니다. 현재 화면의 자료 기준일을 확인하세요.')
    if snapshots:
        choices=sorted(snapshots,key=lambda r:(r['snapshot_date'],r['sequence']),reverse=True)
        chosen=st.selectbox('보관 버전',choices,format_func=lambda s:f'{s["snapshot_date"]} · 버전 {s["sequence"]} · {s["sha256"][:10]}')
        if st.button('선택 버전 관측자료 불러오기'):
            path=ROOT/chosen['observations_path']
            if path.exists():st.session_state['archive_choice']=(chosen['sha256'],read(path))
        loaded=st.session_state.get('archive_choice')
        if loaded and loaded[0]==chosen['sha256']:
            observed=pd.DataFrame(loaded[1]['today']);paged_table(observed,'observed');export(observed,'observed',f'OOS_관측_{chosen["snapshot_date"]}_v{chosen["sequence"]}.csv',chosen['sha256'])
            rejected=pd.DataFrame(loaded[1].get('validation_exclusions',[]))
            if not rejected.empty:
                with st.expander('분석에서 제외된 행 · 원본은 보존됨'):paged_table(rejected,'rejected')
        if st.button('원본 엑셀 다운로드 준비'):
            path=ROOT/chosen['archive_path']
            if path.exists():st.session_state['workbook_download']=(chosen['sha256'],path.read_bytes())
            else:st.error('보관 파일이 없습니다. Actions 저장 결과를 확인해 주세요.')
        ready=st.session_state.get('workbook_download')
        if ready and ready[0]==chosen['sha256']:
            st.download_button('원본 엑셀 저장',ready[1],file_name=f'고장_RAW_{chosen["snapshot_date"]}_v{chosen["sequence"]}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.write(f'기존 이력 파일 {meta.get("legacy_import_files",0)}개 확인 · 날짜별 알람 {meta.get("legacy_dated_entries",0):,}건 · PORT 단독 알람 {meta.get("legacy_port_entries",0):,}건')
    if not meta.get('legacy_import_files',0):st.info('기존 이력 파일을 아직 찾지 못했습니다. 보유한 alarm_history.json을 저장소 루트에 두고 자료 처리 작업을 실행하면 자동 연결됩니다.')
    for note in meta.get('notes',[]):st.warning(note)
    st.caption('기존 alarm_history.json은 수정하지 않습니다. 사본을 별도로 보존하며 날짜·PORT 이력은 발생시각 미검증으로 표시합니다.')

st.caption(f'자료 기준일 {snapshot} · 처리 {meta["generated_at"]} · JINJU FAULT INSIGHT V3')
