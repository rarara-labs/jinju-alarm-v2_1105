import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import test_history

class SyncTest(unittest.TestCase):
    def test_catches_all_skipped_workbook_revisions_and_rerun_is_idempotent(self):
        fixture=test_history.HistoryTest();fixture.setUp()
        try:
            root=fixture.root
            source=Path(__file__).resolve().parents[1]
            for name in ['fault_core.py','build_data.py','sync_history.py']:
                shutil.copy2(source/name,root/name)
            def git(*args):return subprocess.check_output(['git',*args],cwd=root,stderr=subprocess.DEVNULL)
            git('init','-b','main');git('config','user.email','test@example.invalid');git('config','user.name','Test')
            (root/'.github/workflows').mkdir(parents=True)
            (root/'.github/workflows/fault-snapshot-v3.yml').write_text('name: test')
            workbook=root/'26년_진주품질개선팀_고장_RAW_DATA.xlsx'
            fixture.write();shutil.copy2(fixture.file,workbook)
            git('add','.');git('commit','-m','install with baseline')
            fixture.write(alarm='SW Error');shutil.copy2(fixture.file,workbook)
            git('add',workbook.name);git('commit','-m','same day revision')
            fixture.write(day='2026-11-06');shutil.copy2(fixture.file,workbook)
            git('add',workbook.name);git('commit','-m','next day')
            subprocess.run([sys.executable,'sync_history.py'],cwd=root,check=True,capture_output=True)
            manifest=json.loads((root/'data/v3/manifest.json').read_text())
            self.assertEqual(len(manifest['snapshots']),3)
            self.assertEqual(len(json.loads((root/'data/v3/git-sync.json').read_text())['revisions']),3)
            subprocess.run([sys.executable,'sync_history.py'],cwd=root,check=True,capture_output=True)
            self.assertEqual(len(json.loads((root/'data/v3/manifest.json').read_text())['snapshots']),3)
        finally:fixture.tearDown()

if __name__=='__main__':unittest.main()
