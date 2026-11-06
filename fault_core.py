"""Pure workbook processing. No network calls or persistent writes."""
from __future__ import annotations
from collections import defaultdict, Counter
from datetime import datetime, timedelta
from hashlib import sha256
import json
import re
import pandas as pd

PRIORITY = ['Input Power Failure','Link Failure','No Connection','SW Error','FRU General Problem']
BAD = {'', 'nan', 'NaT', 'None', '-', 'Service Unavailable'}
OOS = '5G_LTE OOS_진주'
AUX = ['중계기 및 MIBOS 알람','gREMS','RMS_A망 미복구','RMS_DACS 미복구','RMS_통합RCU미복구']
RAW = ['5G Raw','LTE Raw','3G MOD']
FIELDS = ['시스템','Port','발생시각','장비명','시군구','고장구분','고장구분(중분류)','복구/미복구 상세내역','복구시각','세부알람','5G Link 알람','알람등급','경과일']

def clean(v):
    if v is None: return ''
    try:
        if pd.isna(v): return ''
    except (ValueError, TypeError): pass
    return str(v).strip()

def dt(v):
    if not clean(v): return ''
    # Invalid/missing timestamps never become today's date or an invented event.
    if isinstance(v, (int, float)): return ''
    parsed = pd.to_datetime(v, errors='coerce')
    if pd.isna(parsed): return ''
    if getattr(parsed, 'tzinfo', None): parsed=parsed.tz_convert('Asia/Seoul').tz_localize(None)
    result=parsed.strftime('%Y-%m-%d %H:%M:%S')
    if parsed.microsecond:result+='.'+f'{parsed.microsecond:06d}'.rstrip('0')
    return result

def day(v): return dt(v)[:10]
def valid_alarm(v): return clean(v) not in BAD

def best(values):
    values={clean(v) for v in values if valid_alarm(v)}
    return next((p for p in PRIORITY if p in values), ', '.join(sorted(values)))

def event_id(row):
    payload=[clean(row.get('시스템')), clean(row.get('Port')), dt(row.get('발생시각'))]
    if not all(payload): return ''
    return sha256(json.dumps(payload,ensure_ascii=False).encode()).hexdigest()[:32]

def records(frame):
    return [{str(k):clean(v) for k,v in row.items()} for row in frame.to_dict('records')]

def read_workbook(path, as_of):
    notes=[]
    with pd.ExcelFile(path) as excel:
        sheet_map={name.strip():name for name in excel.sheet_names}
        sheets={}
        for name in [OOS]+RAW+['AAU PORT_0518','RRU PORT_0518']+AUX:
            if name not in sheet_map:
                if name==OOS: raise ValueError('필수 OOS 시트가 없습니다.')
                notes.append(f'누락 시트: {name}')
                sheets[name]=pd.DataFrame()
                continue
            sheets[name]=pd.read_excel(excel,sheet_name=sheet_map[name]).dropna(how='all')
            sheets[name].columns=sheets[name].columns.map(lambda x:str(x).strip())
    oos=sheets[OOS]
    for col in ['본부','시스템','Port','발생시각']:
        if col not in oos: raise ValueError(f'OOS 필수 열 누락: {col}')
    cutoff=(pd.Timestamp(as_of)+pd.Timedelta(days=30)).strftime('%Y-%m-%d')
    rows=[]; rejected=[]
    for number,row in enumerate(oos.to_dict('records'),2):
        row={k:clean(v) for k,v in row.items()}
        row['발생시각']=dt(row.get('발생시각'))
        row['기준일']=day(row.get('본부'))
        row['event_id']=event_id(row)
        if not row['event_id'] or not row['기준일'] or row['기준일']>cutoff:
            rejected.append({'엑셀행':number,'Port':row.get('Port',''),'본부':row.get('본부',''),'발생시각':row.get('발생시각',''),'사유':'식별자/날짜 누락 또는 기준일 범위 초과'})
            continue
        rows.append(row)
    if not rows: raise ValueError('유효한 OOS 자료가 없습니다. 기존 스냅샷을 유지합니다.')
    # Dedup only within a snapshot day; do NOT erase prior observed dates.
    unique={}
    for row in rows: unique[(row['기준일'],row['event_id'])]=row
    rows=list(unique.values())
    latest=max(row['기준일'] for row in rows)
    today=[r for r in rows if r['기준일']==latest]
    dates=sorted({r['기준일'] for r in rows})
    prev=dates[-2] if len(dates)>1 else ''
    previous={r['Port'] for r in rows if r['기준일']==prev}
    if rejected:notes.append(f'OOS 식별/날짜 검증 제외 {len(rejected)}행 (원본은 보관됨)')
    rf_map=defaultdict(set);du_map=defaultdict(set);raw_counts={}; raw_records=[]; raw_invalid=0
    for name in RAW:
        frame=sheets[name];raw_counts[name]=len(frame)
        if frame.empty:continue
        if not {'DU','specificProblem','objectOfReference'}<=set(frame.columns):
            raise ValueError(f'{name} 필수 열 누락: DU/specificProblem/objectOfReference')
        for rawrow in frame.to_dict('records'):
            rec={str(k):clean(v) for k,v in rawrow.items()}
            rec['시스템']={'5G Raw':'5G','LTE Raw':'LTE','3G MOD':'3G'}[name]
            rec['eventTime']=dt(rawrow.get('eventTime'))
            if rec['eventTime'] and rec.get('DU') and rec.get('specificProblem'):
                identity=[rec.get(k,'') for k in ['시스템','DU','eventTime','specificProblem','objectOfReference','eventType']]
                rec['raw_id']=sha256(json.dumps(identity,ensure_ascii=False).encode()).hexdigest()[:32]
                raw_records.append(rec)
            else:raw_invalid+=1
        for du,alarm,reference in frame[['DU','specificProblem','objectOfReference']].itertuples(index=False,name=None):
            du,alarm=clean(du),clean(alarm)
            if not du or not valid_alarm(alarm):continue
            du_map[du].add(alarm)
            match=re.search(r'FieldReplaceableUnit=([^,\s]+)',clean(reference))
            if match:rf_map[(du,match.group(1))].add(alarm)
    cells={}
    for name,cell_col in [('AAU PORT_0518','NRCellDUId'),('RRU PORT_0518','EUtranCellFDDId')]:
        frame=sheets[name]
        if frame.empty:continue
        columns=['NodeId',cell_col,'RF_port']
        if not set(columns)<=set(frame.columns):
            raise ValueError(f'{name} 매핑 열 누락: {columns}')
        for du,cell,rf in frame[columns].itertuples(index=False,name=None):
            du,cell,rf=clean(du),clean(cell),clean(rf)
            if du and cell and rf:cells[cell]=(du,rf)
    matcher=AlarmMatcher(rf_map,du_map,cells)
    raw_dates={name:sorted({day(v) for v in frame.get('eventTime',[]) if day(v)}) for name,frame in sheets.items() if name in RAW}
    # eventTime is original occurrence time, NOT export time; no date-based RAW filtering.
    if raw_invalid:notes.append(f'RAW 집계에서 식별/발생시각 누락 {raw_invalid}행 제외 (원본은 보관됨)')
    if prev=='':notes.append('직전 기준일이 없어 신규 여부는 비교불가입니다.')
    return {'rows':rows,'today':today,'latest':latest,'prev':prev,'previous':previous,'matcher':matcher,
            'raw_records':raw_records,'aux':{k:records(sheets[k]) for k in AUX},'raw_counts':raw_counts,'raw_dates':raw_dates,'notes':notes,'rejected':rejected}

class AlarmMatcher:
    def __init__(self,rf_map,du_map,cells):
        self.rf_map,self.du_map,self.cells=rf_map,du_map,cells
        self.cache={}
        # Only cells with actual alarm candidates can affect partial-match fallback.
        self.active=[(cell,best(rf_map.get(dr,set()))) for cell,dr in cells.items() if rf_map.get(dr)]
    def get(self,port):
        port=clean(port)
        if not port:return '', '미확인'
        if port in self.cache:return self.cache[port]
        direct=self.cells.get(port)
        if direct and self.rf_map.get(direct):
            result=(best(self.rf_map[direct]),'RAW DU/RF 정확일치')
        else:
            result=None
            for cell,alarm in self.active:
                if cell in port or port in cell:
                    result=(alarm,'RAW DU/RF 부분일치 · 검토');break
            if result is None:
                parts=port.rsplit('_',1)
                du=parts[0] if len(parts)==2 and parts[1].isdigit() else port
                alarm=best(self.du_map.get(du,set()))
                result=(alarm,'RAW DU 일치' if alarm else '미확인')
        self.cache[port]=result
        return result

def oos_alarm(row):
    v=clean(row.get('5G Link 알람'))
    if valid_alarm(v):
        v=v.split(':',1)[-1].strip()
        if valid_alarm(v):return v,'OOS Link'
    v=clean(row.get('세부알람'))
    if valid_alarm(v):return v,'OOS 세부알람'
    return '', '미확인'

def classify_work(text):
    text=clean(text)
    if text in BAD or text=='미복구':return ''
    pairs=[('자동복구','자동복구'),('AAU','AAU 교체/수리'),('정류기','정류기'),('SFP','SFP'),('광점퍼','광점퍼'),('광케이블','광케이블'),('PSU','PSU'),('PRU','PRU'),('공중선','공중선'),('차단기','차단기'),('전원','전원'),('교체','기타 교체'),('복구','기타 복구')]
    return next((label for kw,label in pairs if kw in text),'기타')
