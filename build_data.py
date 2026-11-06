"""Append-only workbook archives + occurrence history + compact read-only UI snapshot."""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from datetime import datetime
import gzip
from hashlib import sha256
import json
from pathlib import Path
import tempfile
import time
from zoneinfo import ZoneInfo
from fault_core import read_workbook, clean, valid_alarm, oos_alarm, FIELDS, dt

ROOT=Path(__file__).resolve().parent
SOURCE_NAME='26년_진주품질개선팀_고장_RAW_DATA.xlsx'

def read_json(path,default):
    if not path.exists():return default
    try:
        content=gzip.decompress(path.read_bytes()) if path.suffix=='.gz' else path.read_bytes()
        value=json.loads(content)
    except Exception as exc:raise ValueError(f'기존 저장자료 손상: {path.name}. 덮어쓰지 않습니다.') from exc
    if not isinstance(value,type(default)):raise ValueError(f'기존 저장자료 형식 오류: {path.name}')
    return value

def encode(value,compressed=False):
    content=json.dumps(value,ensure_ascii=False,separators=(',',':'),allow_nan=False).encode()
    return gzip.compress(content,mtime=0) if compressed else content

def atomic_bytes(path,content):
    path.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=path.parent,delete=False) as fp:
        fp.write(content);temp=Path(fp.name)
    temp.replace(path)

def legacy_history(root):
    combined={};copies=[];conflicts=0
    for relative in ['alarm_history.json','data/alarm_history.json','data/previous_alarm_history.json']:
        path=root/relative
        if not path.exists():continue
        data=read_json(path,{})
        for date,entries in data.items():
            if not isinstance(entries,dict):raise ValueError(f'구 이력 형식 오류: {relative}/{date}')
            for port,alarm in entries.items():
                if not isinstance(alarm,str):raise ValueError(f'구 이력 알람 형식 오류: {relative}')
                if not valid_alarm(alarm):continue
                dest=combined.setdefault(date,{})
                if port in dest and dest[port]!=alarm:conflicts+=1
                else:dest[port]=alarm
        copies.append((relative,path.read_bytes()))
    return combined,copies,conflicts

def choose_source(root):
    # Existing daily upload path is kept; no moving/renaming workbook required.
    for relative in [SOURCE_NAME,'data/daily.xlsx']:
        if (root/relative).is_file():return root/relative
    raise ValueError(f'엑셀이 없습니다: {SOURCE_NAME} 또는 data/daily.xlsx')

def build(root=ROOT,source=None,revision='',captured_at=None):
    started=time.perf_counter();root=Path(root);base=root/'data/v3'
    source=Path(source) if source else choose_source(root)
    if not source.is_file():raise ValueError(f'입력 엑셀 없음: {source}')
    captured_at=captured_at or datetime.now(ZoneInfo('Asia/Seoul')).isoformat(timespec='seconds')
    content=source.read_bytes();digest=sha256(content).hexdigest()
    raw_store=read_json(base/'raw_events.json.gz',{'schema_version':3,'events':{}})
    manifest=read_json(base/'manifest.json',{'schema_version':3,'snapshots':[]})
    store=read_json(base/'events.json.gz',{'schema_version':3,'events':{}})
    if manifest.get('schema_version')!=3 or store.get('schema_version')!=3:raise ValueError('지원하지 않는 이력 버전')
    if not isinstance(store.get('events'),dict) or not isinstance(manifest.get('snapshots'),list):raise ValueError('잘못된 누적 이력 구조')
    # All parsing/validation happens BEFORE touching persistent files.
    legacy,copies,conflicts=legacy_history(root)
    book=read_workbook(source,captured_at[:10]);latest=book['latest']
    if conflicts:book['notes'].append(f'기존 이력 충돌 {conflicts}건. 원본별 보존, 선순위 값 유지.')
    raw_events=raw_store['events']
    for row in book['raw_records']:
        rid=row['raw_id'];old=raw_events.get(rid)
        if old is None:
            old=dict(row);old.update({'최초관측일':latest,'최종관측일':latest,'관측일':[]});raw_events[rid]=old
        old['최초관측일']=min(old['최초관측일'],latest);old['최종관측일']=max(old['최종관측일'],latest)
        if latest not in old['관측일']:old['관측일'].append(latest);old['관측일'].sort()
        old['관측일수']=len(old['관측일'])
    events=store['events'];known=next((s for s in manifest['snapshots'] if s['sha256']==digest),None)
    sequence=known['sequence'] if known else len(manifest['snapshots'])+1
    archive_rel=f'archive/workbooks/{latest}/{digest}.xlsx'
    if known:
        for field in ['archive_path','observations_path']:
            if not (root/known[field]).is_file():raise ValueError(f'보관 이력의 파일이 없습니다: {known[field]}')
        if sha256((root/known['archive_path']).read_bytes()).hexdigest()!=digest:raise ValueError('보관 원본이 손상됐습니다. 덮어쓰지 않습니다.')
    observation_rel=f'archive/observations/{latest}/{digest}.json.gz'
    # A legacy day/port mapping is only attached when it identifies one occurrence.
    legacy_candidates=defaultdict(set)
    for r in book['rows']:legacy_candidates[(r['기준일'],r['Port'])].add(r['event_id'])
    now_rows=[]
    for row in sorted(book['rows'],key=lambda r:(r['기준일'],r['발생시각'])):
        eid=row['event_id'];reference=row['기준일'];port=row['Port']
        event=events.get(eid)
        if event is None:
            event={'event_id':eid,'시스템':row['시스템'],'Port':port,'발생시각':row['발생시각'],
                   '최초기준일':reference,'최종기준일':reference,'관측일':[],
                   '최초확인알람':'','최초알람근거':'','최초알람확인일':'','알람종류':[],
                   '알람':'','알람근거':'미확인','revision_sequence':0}
            events[eid]=event
        alarm,alarm_source=('','미확인')
        if reference==latest:alarm,alarm_source=book['matcher'].get(port)
        if not alarm:alarm,alarm_source=oos_alarm(row)
        if not alarm:
            old=legacy.get(reference,{}).get(port,'')
            if old and len(legacy_candidates[(reference,port)])==1:
                alarm,alarm_source=old,'구이력 날짜/PORT · 발생시각 미검증'
        # Never reassign a current RAW or port-only legacy alarm to a past occurrence.
        if alarm:
            if alarm not in event['알람종류']:event['알람종류'].append(alarm)
            if not event['최초확인알람']:
                event['최초확인알람']=alarm;event['최초알람근거']=alarm_source;event['최초알람확인일']=reference
        if reference not in event['관측일']:event['관측일'].append(reference);event['관측일'].sort()
        event['최초기준일']=min(event['최초기준일'],reference)
        if reference>event['최종기준일'] or (reference==event['최종기준일'] and sequence>=event['revision_sequence']):
            for key in FIELDS:event[key]=clean(row.get(key))
            event['최종기준일']=reference;event['revision_sequence']=sequence
            # Missing match is explicit; a previous occurrence-specific value may be retained with its source date.
            if alarm:event['알람']=alarm;event['알람근거']=alarm_source
        event['관측일수']=len(event['관측일'])
        resolved=alarm  # No port-only fallback across occurrences.
        if not resolved and event.get('최초확인알람'):
            resolved=event.get('알람') or event['최초확인알람']
            alarm_source='동일 발생 이력 · 현재 RAW 미확인'
        observation={key:clean(row.get(key)) for key in FIELDS}
        observation.update({'event_id':eid,'기준일':reference,'알람':resolved,'알람근거':alarm_source})
        # Full observations of latest day (including same-day revisions) are immutable.
        if reference==latest:
            checked=bool(row.get('고장구분')) and clean(row.get('고장구분')) not in {'nan','None','-'}
            is_new=(port not in book['previous']) if book['prev'] else None
            observation.update({'신규':is_new,'점검기재':checked,
                '현장상태':'점검·조치' if checked else '비교불가' if is_new is None else '신규' if is_new else '미점검'})
            now_rows.append(observation)
    # Replaying old revisions can grow history but cannot roll the current dashboard backward.
    old_current=read_json(base/'current.json',{})
    old_order=(old_current.get('meta',{}).get('snapshot_date',''),old_current.get('meta',{}).get('sequence',0))
    update_current=(latest,sequence)>=old_order
    active={r['event_id'] for r in (now_rows if update_current else old_current.get('today',[]))}
    for eid,e in events.items():
        explicit=dt(e.get('복구시각'))
        e['상태']='복구' if explicit else '미복구 관측' if eid in active else '금일 미관측 · 복구 추정'
    # Legacy values never enter rankings as verified contemporary observations.
    summary={'today_total':len(now_rows),'today_new':sum(r['신규'] is True for r in now_rows),
        'today_checked':sum(r['점검기재'] for r in now_rows),
        'today_unchecked':sum(r['현장상태']=='미점검' for r in now_rows),
        'raw_event_total':len(raw_events),'event_total':len(events),'snapshot_total':len(manifest['snapshots'])+(0 if known else 1),
        'source_workbook_oos_rows':len(book['rows']),
        'matched_today':sum(bool(r['알람']) for r in now_rows),
        'repeated_ports':sum(n>1 for n in Counter((e['시스템'],e['Port']) for e in events.values()).values())}
    current={'meta':{'schema_version':3,'snapshot_date':latest,'previous_date':book['prev'],
         'generated_at':captured_at,'source_sha256':digest,'source_revision':revision,'sequence':sequence,
         'archive_path':archive_rel,'observations_path':observation_rel,
         'build_seconds':round(time.perf_counter()-started,3),'notes':book['notes'],
         'raw_counts':book['raw_counts'],'raw_occurrence_dates':book['raw_dates'],
         'rejected_rows':len(book['rejected']),
         'legacy_import_files':len(copies),'legacy_dated_entries':sum(len(v) for k,v in legacy.items() if k!='_ports'),
         'legacy_port_entries':len(legacy.get('_ports',{}))},'summary':summary,'today':now_rows,'sheets':book['aux']}
    snapshot={'sequence':sequence,'sha256':digest,'snapshot_date':latest,'captured_at':captured_at,
              'revision':revision,'archive_path':archive_rel,'observations_path':observation_rel,
              'today_count':len(now_rows),'raw_counts':book['raw_counts']}
    writes={}
    if not known:
        if (root/archive_rel).exists() and (root/archive_rel).read_bytes()!=content:raise ValueError('보관 원본 해시 충돌')
        writes[root/archive_rel]=content
        writes[root/observation_rel]=encode({'meta':current['meta'],'today':now_rows,'validation_exclusions':book['rejected']},True)
        manifest['snapshots'].append(snapshot)
    # Preserve exact old legacy files, including entries not found in this workbook.
    for relative,raw in copies:
        h=sha256(raw).hexdigest();path=root/f'archive/legacy/{h}.json'
        if not path.exists():writes[path]=raw
    writes[base/'manifest.json']=encode(manifest)
    writes[base/'events.json.gz']=encode(store,True)
    writes[base/'raw_events.json.gz']=encode(raw_store,True)
    if update_current:writes[base/'current.json']=encode(current)
    elif old_current:
        old_current['summary']['event_total']=len(events)
        old_current['summary']['raw_event_total']=len(raw_events)
        old_current['summary']['snapshot_total']=len(manifest['snapshots'])
        old_current['summary']['repeated_ports']=summary['repeated_ports']
        writes[base/'current.json']=encode(old_current)
    # Stage a validated complete file set; current.json is replaced last.
    for path,raw in writes.items():atomic_bytes(path,raw)
    return {'snapshot_date':latest,'events':len(events),'today':len(now_rows),'new_snapshot':not known,
            'sha256':digest,'elapsed_seconds':round(time.perf_counter()-started,3),'current_bytes':len(encode(current)),'notes':book['notes']}

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--source');parser.add_argument('--root',default=str(ROOT));parser.add_argument('--revision',default='')
    args=parser.parse_args()
    print(json.dumps(build(args.root,args.source,args.revision),ensure_ascii=False,indent=2))
