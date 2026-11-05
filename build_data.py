# -*- coding: utf-8 -*-
"""일일 Excel을 경량 JSON으로 가공하고 날짜별 알람 이력을 비파괴 병합한다.
Run: python build_data.py
"""
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from datetime import datetime, date
import pandas as pd

ROOT=Path(__file__).resolve().parent
FILE=ROOT/'data/daily.xlsx'
OUT=ROOT/'data/processed.json'
HIST=ROOT/'data/alarm_history.json'
LEGACY=ROOT/'data/previous_alarm_history.json'  # optional V1 history import

PRIORITY=['Input Power Failure','Link Failure','No Connection','SW Error','FRU General Problem']
BAD={'','nan','None','Service Unavailable','NaT','-'}

def clean(v):
    if v is None: return ''
    try:
        if pd.isna(v): return ''
    except (TypeError,ValueError): pass
    if isinstance(v,(pd.Timestamp,datetime,date)): return str(v)[:19]
    return str(v).strip()

def datekey(v):
    dt=pd.to_datetime(v,errors='coerce')
    return dt.strftime('%Y-%m-%d') if pd.notna(dt) else ''

def best_alarm(alarms):
    al=[a for a in alarms if a and a not in BAD]
    for kw in PRIORITY:
        if kw in al:return kw
    return ', '.join(sorted(set(al))) if al else ''

def load_sheet(excel,name):
    try:
        df=pd.read_excel(excel,sheet_name=name)
        df.columns=df.columns.map(lambda s:str(s).strip())
        return df.dropna(how='all')
    except ValueError:
        print('WARN: Missing sheet:',name)
        return pd.DataFrame()

def rowrecords(df):
    # Convert every Excel-ish value to JSON-compatible short text.
    return [{str(k):clean(v) for k,v in r.items()} for r in df.to_dict(orient='records')]

def raw_alarm_mapping(excel):
    rf_map=defaultdict(set); du_map=defaultdict(set)
    for name in ['5G Raw','LTE Raw','3G MOD ']:
        df=load_sheet(excel,name)
        if df.empty: continue
        for du,sp,obj in df[['DU','specificProblem','objectOfReference']].itertuples(index=False,name=None):
            du=clean(du); sp=clean(sp)
            if not du or sp in BAD: continue
            du_map[du].add(sp)
            m=re.search(r'FieldReplaceableUnit=([^,\s]+)',clean(obj))
            if m:rf_map[(du,m.group(1).strip())].add(sp)
    cellmap={}
    # Keep positional mapping compatible with V1. Some AAU exports have
    # different columns, so mapping falls back to DU match if needed.
    for sheet in ('AAU PORT_0518','RRU PORT_0518'):
        df=load_sheet(excel,sheet)
        if df.empty or len(df.columns)<6:continue
        for row in df.itertuples(index=False,name=None):
            du=clean(row[3]); cell=clean(row[4]); rf=clean(row[5])
            if du and rf and cell and cell not in BAD:
                cellmap[cell]=(du,rf)
    return rf_map,du_map,cellmap

def get_alarm_for_port(port,rf_map,du_map,cellmap,cache):
    port=clean(port)
    if port in cache: return cache[port]
    cand=None
    if port in cellmap: cand=cellmap[port]
    if cand is None:
        for cell,dr in cellmap.items():
            if cell and port and (cell in port or port in cell):
                cand=dr;break
    if cand:
        vals=rf_map.get(cand,set())
        if vals:
            res=best_alarm(vals);cache[port]=res;return res
    du=port.rsplit('_',1)[0] if '_' in port and port.rsplit('_',1)[-1].isdigit() else port
    res=best_alarm(du_map.get(du,set()))
    cache[port]=res
    return res

def oos_alarm(row):
    lv=clean(row.get('5G Link 알람'))
    if lv and lv not in BAD:
        return (lv.split(':',1)[-1].strip() if ':' in lv else lv),'OOS Link'
    sp=clean(row.get('세부알람'))
    if sp and sp not in BAD:return sp,'OOS 세부알람'
    return '',''

def classify_work(t):
    s=clean(t)
    if not s or s in BAD or s=='미복구':return ''
    for kw,label in [('자동복구','자동복구'),('AAU','AAU 교체/수리'),('정류기','정류기'),('SFP','SFP'),('광점퍼','광점퍼'),('광케이블','광케이블'),('PSU','PSU'),('PRU','PRU'),('공중선','공중선'),('차단기','차단기'),('전원','전원'),('교체','기타 교체'),('복구','기타 복구')]:
        if kw in s:return label
    return '기타'

def main():
    if not FILE.exists():raise SystemExit('Missing daily workbook: data/daily.xlsx')
    xls=pd.ExcelFile(FILE)
    oos=load_sheet(xls,'5G_LTE OOS_진주')
    if oos.empty:raise SystemExit('OOS worksheet is empty')
    for required in ['본부','Port','발생시각']:
        if required not in oos.columns:raise SystemExit(f'OOS missing mandatory column: {required}')
    oos['기준일']=oos['본부'].map(datekey)
    oos['Port']=oos['Port'].map(clean)
    oos['발생시각']=oos['발생시각'].map(clean)
    oos=oos.drop_duplicates(['Port','발생시각'],keep='last').copy()
    dates=sorted(d for d in oos['기준일'].unique() if d)
    if not dates:raise SystemExit('No valid snapshot date in OOS')
    latest=dates[-1]
    today=oos[oos['기준일']==latest].copy()
    prev=dates[-2] if len(dates)>1 else None
    prev_ports=set(oos.loc[oos['기준일']==prev,'Port']) if prev else set()
    rf_map,du_map,cell_map=raw_alarm_mapping(xls)
    cache={}
    # Preserve old data! Existing history JSON may be copied from V1.
    if HIST.exists():
        try: history=json.loads(HIST.read_text('utf-8'))
        except (json.JSONDecodeError,UnicodeDecodeError):
            raise SystemExit('History JSON invalid - refusing to overwrite')
        if not isinstance(history,dict): raise SystemExit('History JSON not an object - refusing to overwrite')
    else:history={}
    history.setdefault('_ports',{})
    # Optional previous V1 archive takes priority for old dated records.
    # Merge rather than replacing new history, so new V2 records remain safe.
    if LEGACY.exists():
        legacy=json.loads(LEGACY.read_text('utf-8'))
        if not isinstance(legacy,dict):raise SystemExit('Invalid previous history archive')
        for dk, entries in legacy.items():
            if not isinstance(entries,dict):continue
            existing=history.setdefault(dk,{})
            for port, alarm in entries.items():
                if alarm:existing[str(port)]=alarm
    date_update_count=0
    # Populate historical dates from OOS only if missing; never
    # reassign CURRENT raw alarm to a past date.
    for _,r in oos.iterrows():
        dk=r['기준일'];port=r['Port']
        if not dk or not port or dk==latest:continue  # latest day: RAW must take priority
        al,_=oos_alarm(r)
        if al and not history.setdefault(dk,{}).get(port):
            history[dk][port]=al;date_update_count+=1
    records=[]
    today_ports=set(today['Port'])
    for _,r in today.iterrows():
        row=r.to_dict();port=clean(row.get('Port'))
        raw=get_alarm_for_port(port,rf_map,du_map,cell_map,cache)
        oos_val,oos_source=oos_alarm(row)
        original=history.get(latest,{}).get(port,'')
        alarm=raw or oos_val or original or history['_ports'].get(port,'')
        source=('RAW' if raw else oos_source if oos_val else '저장이력' if alarm else '미확인')
        # Preserve dated history when already present. Never overwrite known value.
        if alarm and not original:
            history.setdefault(latest,{})[port]=alarm;date_update_count+=1
        # Latest port lookup may update, but NEVER use it to replace dated history.
        if alarm: history['_ports'][port]=alarm
        gj=clean(row.get('고장구분'))
        state='점검·조치' if gj and gj not in BAD else '신규' if port not in prev_ports else '미점검'
        records.append({key:clean(row.get(key)) for key in ['시스템','장비명','Port','시군구','발생시각','고장구분','복구/미복구 상세내역']}
                       |{'알람':alarm,'알람근거':source,'현장상태':state})
    # Generate historical view: never mutate old dated alarm values.
    counts=Counter(oos['Port'])
    all_rows=[]
    for _,r in oos.sort_values('발생시각',ascending=False).iterrows():
        row=r.to_dict();dk=row['기준일'];port=row['Port']
        prior=history.get(dk,{}).get(port,'')
        oos_val,_=oos_alarm(row)
        alarm=prior or oos_val or history['_ports'].get(port,'') or ''
        out={key:clean(row.get(key)) for key in ['본부','발생시각','장비명','Port','시스템','시군구','고장구분','복구/미복구 상세내역']}
        out.update({'알람':alarm,'다발횟수':counts[port]})
        all_rows.append(out)
    trends=(oos.groupby('기준일').size().sort_index().tail(21))
    recurrent=(oos.groupby('Port',dropna=False).agg(횟수=('발생시각','size'),최근발생=('발생시각','max'),장비명=('장비명','last')).reset_index())
    recurrent=recurrent[recurrent['횟수']>1].sort_values('횟수',ascending=False).head(40)
    work_count=Counter(filter(None,(classify_work(t) for t in oos.get('복구/미복구 상세내역',[]))))
    sheets={}
    for name in ['13시 알람 공유','중계기 및 MIBOS 알람','gREMS','RMS_A망 미복구','RMS_DACS 미복구','RMS_통합RCU미복구']:
        sheets[name]=rowrecords(load_sheet(xls,name))
    usable_days=[k for k,v in history.items() if k!='_ports' and isinstance(v,dict) and v]
    summary={'today_total':len(records),'today_new':sum(r['현장상태']=='신규' for r in records),
             'today_checked':sum(r['현장상태']=='점검·조치' for r in records),
             'today_unchecked':sum(r['현장상태']=='미점검' for r in records),
             'all_total':len(oos),'repeated_ports':sum(n>1 for n in counts.values()),'history_days':len(usable_days)}
    result={'meta':{'snapshot_date':latest,'generated_at':datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'raw_count':sum(len(x) for x in [rf_map,du_map]),'added_history_entries':date_update_count},
            'summary':summary,'today':records,'all_oos':all_rows,
            'trends':[{'날짜':date,'건수':int(count)} for date,count in trends.items()],
            'repeated':rowrecords(recurrent),'work_types':[{'작업유형':k,'건수':v} for k,v in work_count.most_common()],
            'sheets':sheets}
    OUT.write_text(json.dumps(result,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
    HIST.write_text(json.dumps(history,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'OK {latest}: {len(records)} today / {len(oos)} total / {len(usable_days)} history days / {date_update_count} added; JSON {OUT.stat().st_size//1024} KiB')

if __name__=='__main__':main()
