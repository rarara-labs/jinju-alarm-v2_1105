"""Meaningful regression tests for archival safety and counting semantics."""
import copy
import gzip
import json
from pathlib import Path
import sys
import tempfile
import unittest
import pandas as pd
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from build_data import build, read_json
from fault_core import AlarmMatcher, event_id

class HistoryTest(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
        self.file=self.root/'input.xlsx'
    def tearDown(self):self.temp.cleanup()
    def write(self,day='2026-11-05',occur='2026-11-05 09:00:00',alarm='Input Power Failure',port='du_01',system='LTE',historical=None,raw=True,extra=None,missing_time=False):
        rows=[{'본부':day,'시스템':system,'Port':port,'발생시각':'' if missing_time else occur,'장비명':'검증국소','시군구':'진주','고장구분':'','세부알람':'Service Unavailable','복구/미복구 상세내역':''}]
        if historical:rows=historical+rows
        if extra:rows+=extra
        rawrows=[{'DU':'du','eventTime':occur,'specificProblem':alarm,'objectOfReference':'FieldReplaceableUnit=RF1','eventType':'ProcessingErrorAlarm'}] if raw else []
        with pd.ExcelWriter(self.file) as w:
            pd.DataFrame(rows).to_excel(w,sheet_name='5G_LTE OOS_진주',index=False)
            for name in ['5G Raw','LTE Raw','3G MOD ']:
                pd.DataFrame(rawrows,columns=['DU','eventTime','specificProblem','objectOfReference','eventType']).to_excel(w,sheet_name=name,index=False)
            pd.DataFrame([{'NodeId':'du','NRCellDUId':'du_01','RF_port':'RF1'}]).to_excel(w,sheet_name='AAU PORT_0518',index=False)
            pd.DataFrame([{'NodeId':'du','EUtranCellFDDId':'du_01','RF_port':'RF1'}]).to_excel(w,sheet_name='RRU PORT_0518',index=False)
    def run_build(self):return build(self.root,self.file,captured_at='2026-11-06T12:00:00+09:00')
    def events(self):return read_json(self.root/'data/v3/events.json.gz',{})['events']
    def manifest(self):return read_json(self.root/'data/v3/manifest.json',{})['snapshots']
    def test_daily_persistence_is_one_event_and_two_observations(self):
        self.write();self.run_build();first=self.file.read_bytes()
        self.write(day='2026-11-06');self.run_build()
        self.assertEqual(len(self.events()),1)
        e=next(iter(self.events().values()));self.assertEqual(e['관측일수'],2)
        self.assertEqual(len(self.manifest()),2)
        self.assertEqual((self.root/self.manifest()[0]['archive_path']).read_bytes(),first)
        for snapshot in self.manifest():self.assertTrue((self.root/snapshot['observations_path']).exists())
    def test_repeated_occurrence_is_separate(self):
        self.write();self.run_build();self.write(day='2026-11-06',occur='2026-11-06 09:00:00');self.run_build()
        self.assertEqual(len(self.events()),2)
    def test_same_day_revision_keeps_both_and_first_alarm(self):
        self.write(alarm='Link Failure');self.run_build();self.write(alarm='SW Error');self.run_build()
        self.assertEqual(len(self.manifest()),2)
        e=next(iter(self.events().values()));self.assertEqual(e['최초확인알람'],'Link Failure');self.assertEqual(e['알람'],'SW Error')
        self.assertEqual(e['관측일수'],1)
    def test_identical_workbook_is_idempotent(self):
        self.write();a=self.run_build();b=self.run_build()
        self.assertEqual(a['sha256'],b['sha256']);self.assertEqual(len(self.manifest()),1);self.assertFalse(b['new_snapshot'])
    def test_current_raw_never_assigned_to_old_occurrence(self):
        old={'본부':'2026-11-04','시스템':'LTE','Port':'du_01','발생시각':'2026-11-04 08:00:00','세부알람':'Service Unavailable'}
        self.write(historical=[old]);self.run_build()
        events=sorted(self.events().values(),key=lambda e:e['발생시각'])
        self.assertEqual(events[0]['최초확인알람'],'');self.assertEqual(events[1]['최초확인알람'],'Input Power Failure')
    def test_legacy_preserved_not_overwritten(self):
        legacy={'2026-11-04':{'du_01':'Link Failure'},'_ports':{'du_01':'SW Error'}}
        path=self.root/'alarm_history.json';path.write_text(json.dumps(legacy));original=path.read_bytes()
        old={'본부':'2026-11-04','시스템':'LTE','Port':'du_01','발생시각':'2026-11-04 08:00:00'}
        self.write(historical=[old]);self.run_build()
        self.assertEqual(path.read_bytes(),original)
        oldevent=min(self.events().values(),key=lambda e:e['발생시각'])
        self.assertEqual(oldevent['최초확인알람'],'Link Failure');self.assertTrue(oldevent['최초알람근거'].startswith('구이력'))
        self.assertTrue(list((self.root/'archive/legacy').glob('*.json')))
    def test_port_only_legacy_not_assigned_to_new_fault(self):
        (self.root/'alarm_history.json').write_text(json.dumps({'_ports':{'du_01':'SW Error'}}))
        self.write(raw=False);self.run_build();self.assertEqual(next(iter(self.events().values()))['최초확인알람'],'')
    def test_ambiguous_legacy_not_assigned(self):
        legacy={'2026-11-05':{'du_01':'Link Failure'}}
        (self.root/'alarm_history.json').write_text(json.dumps(legacy))
        self.write(raw=False,extra=[{'본부':'2026-11-05','시스템':'LTE','Port':'du_01','발생시각':'2026-11-05 10:00:00'}]);self.run_build()
        self.assertTrue(all(not e['최초확인알람'] for e in self.events().values()))
    def test_corrupt_store_refuses_overwrite(self):
        self.write();self.run_build();path=self.root/'data/v3/events.json.gz';path.write_bytes(b'corrupt')
        oldmanifest=(self.root/'data/v3/manifest.json').read_bytes()
        with self.assertRaises(ValueError):self.run_build()
        self.assertEqual(path.read_bytes(),b'corrupt');self.assertEqual((self.root/'data/v3/manifest.json').read_bytes(),oldmanifest)
    def test_malformed_legacy_refuses_before_writes(self):
        self.write();(self.root/'alarm_history.json').write_text('{invalid')
        with self.assertRaises(ValueError):self.run_build()
        self.assertFalse((self.root/'data/v3').exists())
    def test_missing_timestamp_fails_without_invented_event(self):
        self.write(missing_time=True)
        with self.assertRaises(ValueError):self.run_build()
        self.assertFalse((self.root/'archive').exists())
    def test_old_replay_does_not_roll_current_backward(self):
        self.write(day='2026-11-06',alarm='SW Error');self.run_build()
        self.write(day='2026-11-05',alarm='Link Failure');self.run_build()
        current=read_json(self.root/'data/v3/current.json',{})
        self.assertEqual(current['meta']['snapshot_date'],'2026-11-06')
        e=next(iter(self.events().values()));self.assertEqual(e['알람'],'SW Error')
    def test_raw_duplicate_observations_deduplicate(self):
        self.write();self.run_build();self.write(day='2026-11-06');self.run_build()
        raw=read_json(self.root/'data/v3/raw_events.json.gz',{})['events']
        self.assertEqual(len(raw),3) # Same record exported in each of 5G/LTE/3G remains system-specific.
        self.assertTrue(all(len(r['관측일'])==2 for r in raw.values()))
    def test_aau_uses_named_columns_not_obsolete_positions(self):
        self.write(system='5G',port='nr-cell-1',raw=True)
        with pd.ExcelFile(self.file) as x:frames={n:pd.read_excel(x,n) for n in x.sheet_names}
        frames['AAU PORT_0518']=pd.DataFrame([{'gNBId':'x','Cell':'x','cellLocalId':'x','Unnamed: 3':'wrong','duh_eqp_nm':'wrong','NodeId':'du','NRCellDUId':'nr-cell-1','RF_port':'RF1'}])
        with pd.ExcelWriter(self.file) as w:
            for n,f in frames.items():f.to_excel(w,sheet_name=n,index=False)
        self.run_build();e=next(iter(self.events().values()));self.assertEqual(e['최초알람근거'],'RAW DU/RF 정확일치')
    def test_subsecond_occurrences_are_not_collapsed(self):
        a={'시스템':'LTE','Port':'du_01','발생시각':'2026-11-05 09:00:00.123'}
        b=dict(a,발생시각='2026-11-05 09:00:00.456')
        self.assertNotEqual(event_id(a),event_id(b))
    def test_corrupt_archive_refuses_overwrite(self):
        self.write();self.run_build()
        archive=self.root/self.manifest()[0]['archive_path'];archive.write_bytes(b'bad')
        with self.assertRaises(ValueError):self.run_build()
        self.assertEqual(archive.read_bytes(),b'bad')
    def test_partial_fallback_skips_candidates_without_alarm(self):
        matcher=AlarmMatcher({('du','good'):{'Link Failure'}},{},{'p':('du','empty'),'port':('du','good')})
        self.assertEqual(matcher.get('ports')[0],'Link Failure')

if __name__=='__main__':unittest.main()
