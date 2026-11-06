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
st.set_page_config(page_title='JINJU · 고장 현황',page_icon='📡',layout='wide',initial_sidebar_state='expanded')
CSS='''<style>
:root{--j-teal:#087e83;--j-ink:#172b3b;--j-sub:#72828f;--j-line:#e3eaef}
[data-testid="stAppViewContainer"]{background:#f3f6f8;color:#172b3b}
[data-testid="stHeader"]{background:transparent}
html,body,[class*="css"]{font-family:system-ui,-apple-system,"Apple SD Gothic Neo","Malgun Gothic",sans-serif}
.block-container{padding:2rem 2rem 3rem!important;max-width:1380px}
[data-testid="stSidebar"]{background:#fff;border-right:1px solid #e3eaef;min-width:210px;max-width:245px}
[data-testid="stSidebar"] [data-testid="stRadio"] label{padding:9px 4px}
[data-testid="stSidebar"] [data-testid="stRadio"] label p{font-size:14px}
.j-brand{font-size:22px;font-weight:800;letter-spacing:-.8px;color:#087e83;padding:2px 0 24px}.j-brand small{display:block;font-size:10px;letter-spacing:2px;font-weight:600;color:#72828f;margin-top:4px}
.j-crumb{font-size:11px;letter-spacing:1px;color:#72828f;margin-bottom:20px}.j-title{font-size:29px!important;line-height:1.3!important;font-weight:800;letter-spacing:-1px;margin:0!important;padding:0!important;color:#172b3b}.j-subtitle{font-size:13px;color:#72828f;margin:8px 0 22px}
.j-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:6px 0 22px}.j-kpi{background:#fff;border:1px solid #e3eaef;border-radius:13px;padding:19px 21px}.j-kpi .label{font-size:12px;color:#72828f}.j-kpi .value{font-size:34px;letter-spacing:-1px;font-weight:750;margin:7px 0 5px;line-height:1.2}.j-kpi .hint{font-size:11px;color:#72828f}.j-kpi.focus{background:#087e83;border-color:#087e83;color:#fff}.j-kpi.focus .label,.j-kpi.focus .hint{color:#d5eeed}
.j-heading{display:flex;justify-content:space-between;align-items:center;margin:12px 0 15px;font-size:16px;font-weight:750;letter-spacing:-.3px}.j-heading small{font-size:11px;font-weight:400;color:#72828f}
.j-alert{display:flex;align-items:center;gap:12px;padding:15px 0;border-top:1px solid #e3eaef}.j-alert .num{font-size:11px;color:#72828f}.j-alert .name{font-size:13px;font-weight:700}.j-alert .detail{font-size:11px;color:#72828f;margin-top:4px}.j-tag{display:inline-block;padding:4px 8px;border-radius:6px;font-size:11px;background:#e7f4f3;color:#087e83;white-space:nowrap}.j-tag.warn{background:#fff3df;color:#94651e}.j-tag.new{background:#ffedf1;color:#be3d55}.j-alert .j-tag{margin-left:auto}
.j-table{width:100%;border-collapse:collapse;font-size:12px;text-align:left;background:white}.j-table th{background:#f5f8fa;color:#72828f;font-size:11px;font-weight:500;padding:12px 10px;white-space:nowrap}.j-table td{padding:14px 10px;border-bottom:1px solid #e7edf1;vertical-align:middle}.j-table .device{font-size:12px;font-weight:650;color:#172b3b}.j-table .port{font-size:10px;color:#72828f;margin-top:4px;overflow-wrap:anywhere}.j-table .system{font-size:10px;color:#087e83;font-weight:700;margin-right:5px}.j-table .source{font-size:10px;color:#72828f;margin-top:4px}.j-tablewrap{overflow-x:auto}.j-table tr:last-child td{border-bottom:0}
[data-testid="stVerticalBlockBorderWrapper"]>div{border-color:#e3eaef!important;border-radius:13px!important;background:white}
[data-testid="stDataFrame"]{border-radius:10px;overflow:hidden}[data-testid="stPlotlyChart"]{background:white;border-radius:12px}
.stButton button,.stDownloadButton button{border-radius:8px;border-color:#dbe5ea;font-size:12px;font-weight:600;min-height:40px}
.stDownloadButton button{background:#e7f4f3;color:#087e83}
[data-baseweb="input"],[data-baseweb="select"]>div{background:white;border-radius:8px}
[data-testid="stCaptionContainer"] p{font-size:12px;color:#72828f}
@media(max-width:700px){.block-container{padding:1.2rem .8rem 2rem!important}.j-title{font-size:24px!important}.j-kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.j-kpi{padding:14px}.j-kpi .value{font-size:28px}.j-kpi .hint{font-size:10px}.j-table th,.j-table td{padding:11px 7px}.j-heading{font-size:15px}.j-crumb{font-size:10px;margin-top:8px}.j-alert .name{font-size:12px}}
</style>'''
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
    fig.update_layout(height=max(240,min(440,45*len(frame))) if horizontal else 240,paper_bgcolor='white',plot_bgcolor='white',margin=dict(l=6,r=25,t=10,b=10),font=dict(family='system-ui, sans-serif',color='#536775',size=11),showlegend=False,xaxis_title='',yaxis_title='')
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

def alarm_table(rows):
    if not rows:st.info('표시할 알람이 없습니다.');return
    parts=[]
    for row in rows[:50]:
        state=row.get('현장상태','')
        cls='warn' if state=='미점검' else 'new' if state=='신규' else ''
        parts.append(f'<tr><td><span class="system">{esc(row.get("시스템"))}</span><span class="device">{esc(row.get("장비명"))}</span><div class="port">{esc(row.get("Port"))}</div></td><td>{esc(row.get("알람") or "확인필요")}<div class="source">{esc(row.get("알람근거"))}</div></td><td>{esc(row.get("시군구"))}</td><td><span class="j-tag {cls}">{esc(state)}</span><div class="source">{esc(row.get("고장구분"))}</div></td></tr>')
    st.markdown('<div class="j-tablewrap"><table class="j-table"><thead><tr><th>시스템 / 장비</th><th>알람 / 근거</th><th>지역</th><th>조치 상태</th></tr></thead><tbody>'+''.join(parts)+'</tbody></table></div>',unsafe_allow_html=True)
    if len(rows)>50:st.caption('요약 표는 50건까지 표시합니다. 아래 상세 조회를 이용하세요.')

def history(kind='oos'):
    path=DATA/('events.json.gz' if kind=='oos' else 'raw_events.json.gz')
    if not path.exists():st.warning('누적 자료가 아직 없습니다. 자료 처리 작업을 먼저 실행해 주세요.');return pd.DataFrame()
    data=read(path)
    return pd.DataFrame(list(data['events'].values()))

with st.sidebar:
    st.markdown('<div class="j-brand">◉ JINJU<small>FAULT INSIGHT</small></div>',unsafe_allow_html=True)
    menus=['통합 현황','금일 알람','중계기 · MIBOS','기타 알람','OOS 이력','보고서 · 알람 분석','원본 · 저장 상태']
    if '_next_menu' in st.session_state:st.session_state['nav']=st.session_state.pop('_next_menu')
    menu=st.radio('조회 메뉴',menus,key='nav',label_visibility='collapsed')
    st.divider();st.caption('진주품질개선팀\n\n매일 갱신되는 고장 현황')
    if st.button('↻ 최신 자료 확인',use_container_width=True):
        read_data.clear();st.rerun()

if not (DATA/'current.json').exists():
    st.title('새 대시보드 자료를 준비하고 있습니다')
    st.info('GitHub Actions에서 Preserve fault history V3가 성공하면 화면이 표시됩니다. 기존 엑셀은 같은 이름으로 계속 교체하면 됩니다.')
    st.stop()
try:db=read(DATA/'current.json')
except Exception as exc:st.error(f'자료 읽기 실패: {exc}');st.stop()
meta=db['meta'];summary=db['summary'];today=db['today'];snapshot=meta['snapshot_date']
st.markdown(f'<div class="j-crumb">NETWORK OPERATIONS / {esc(menu)} &nbsp; · &nbsp; 자료 기준일 {esc(snapshot)}</div>',unsafe_allow_html=True)
titles={'통합 현황':('오늘의 고장 현황','지금 확인할 알람과 조치 현황을 한눈에.'),'금일 알람':('금일 알람','장비와 알람을 빠르게 찾고 조치 상태를 확인하세요.'),'중계기 · MIBOS':('중계기 · MIBOS','LEGACY · MIBOS · gREMS 통합 조회'),'기타 알람':('기타 알람','RMS A망 · DACS · 통합RCU 미복구 현황'),'OOS 이력':('OOS 이력','고장 발생과 날짜별 관측을 구분해 기록합니다.'),'보고서 · 알람 분석':('보고서 · 알람 분석','기간별 대표 알람과 반복 고장을 확인하세요.'),'원본 · 저장 상태':('원본 · 저장 상태','일별 원본과 같은 날 재업로드된 자료까지 보관합니다.')}
title,subtitle=titles[menu]
st.markdown(f'<h1 class="j-title">{esc(title)}</h1><div class="j-subtitle">{esc(subtitle)}</div>',unsafe_allow_html=True)
if menu in ['통합 현황','금일 알람']:
    kpis([('전체 알람',summary['today_total'],'최신 기준일 OOS',False),('신규 PORT',summary['today_new'],'직전 기준일에 없던 PORT',False),('점검 · 조치 기재',summary['today_checked'],'고장구분에 내용이 있는 항목',False),('미점검',summary['today_unchecked'],'전일 존재 · 고장구분 미기재',True)])
    st.caption('신규 PORT와 점검·조치 기재는 중복될 수 있습니다. 신규는 고장 재발 횟수와 다릅니다.')

if menu=='통합 현황':
    left,right=st.columns([1.35,1])
    with left:
        with st.container(border=True):
            heading('먼저 확인할 미점검',f'{summary["today_unchecked"]}건')
            priority=[r for r in today if r['현장상태']=='미점검']
            for n,r in enumerate(priority[:5],1):
                st.markdown(f'<div class="j-alert"><span class="num">{n:02}</span><div><div class="name">{esc(r["장비명"])}</div><div class="detail">{esc(r["시스템"])} · {esc(r["알람"] or "확인필요")}</div></div><span class="j-tag warn">미점검</span></div>',unsafe_allow_html=True)
            if not priority:st.success('미점검 항목이 없습니다.')
            if st.button('미점검 전체 조회 →'):
                st.session_state['_next_menu']='금일 알람';st.session_state['today_state']='미점검';st.rerun()
    with right:
        with st.container(border=True):
            heading('금일 알람 분포','대표 알람 기준')
            import plotly.express as px
            f=pd.Series([r['알람'] or '확인필요' for r in today]).value_counts().rename_axis('알람').reset_index(name='건수')
            if not f.empty:
                fig=px.pie(f,names='알람',values='건수',hole=.72,color_discrete_sequence=['#087e83','#5484b8','#b69455','#8ba4aa','#af7f95'])
                fig.update_traces(textinfo='none');fig.update_layout(height=215,margin=dict(t=5,b=10,l=0,r=0),paper_bgcolor='white',font=dict(size=11,color='#536775'),legend=dict(orientation='h',y=-.1),annotations=[dict(text=f'<b>{len(today)}</b><br>전체 알람',x=.5,y=.5,showarrow=False,font=dict(size=18,color='#172b3b'))])
                st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
    with st.container(border=True):
        heading('금일 알람 살펴보기',f'{len(today)}건')
        alarm_table(today[:8])
    a,b=st.columns(2)
    with a:
        with st.container(border=True):
            heading('시스템별 금일 고장')
            f=pd.Series([r['시스템'] for r in today]).value_counts().rename_axis('시스템').reset_index(name='건수')
            chart(f,'시스템','건수',True)
    with b:
        with st.container(border=True):
            heading('누적 저장 현황')
            st.metric('고유 OOS 발생',f'{summary["event_total"]:,}건')
            st.caption(f'원본 {summary["snapshot_total"]:,}개 버전 보관 · 같은 날 수정본 포함')
            st.caption(f'누적 RAW 기록 {summary.get("raw_event_total",0):,}건 · 중복 관측 제외')
            st.caption('OOS 이력과 보고서 메뉴에서 상세 기간을 선택할 수 있습니다.')

elif menu=='금일 알람':
    a,b,c=st.columns([3,1,1])
    q=a.text_input('장비 · PORT · 알람 검색',placeholder='장비명 · PORT · 지역 · 알람')
    system=b.selectbox('시스템',['전체','5G','LTE'])
    state=c.selectbox('현장 상태',['전체','신규','미점검','점검·조치','비교불가'],key='today_state')
    rows=[r for r in today if (system=='전체' or r['시스템']==system) and (state=='전체' or r['현장상태']==state) and (not q or q.lower() in ' '.join(str(v) for v in r.values()).lower())]
    with st.container(border=True):heading('알람 조회',f'{len(rows)}건');alarm_table(rows)
    frame=pd.DataFrame(rows)
    if not frame.empty:
        with st.expander('발생시각 · PORT · 조치내용 상세'):
            paged_table(frame,'today_detail',['발생시각','시스템','장비명','Port','시군구','알람','알람근거','고장구분','복구/미복구 상세내역'])
        export(frame.drop(columns=['event_id'],errors='ignore'),'today',f'금일알람_{snapshot}.csv',(q,system,state))

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
