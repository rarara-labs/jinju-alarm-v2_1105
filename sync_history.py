"""Catch up every input revision since installation, even when workflow jobs are skipped."""
from pathlib import Path
import json
import subprocess
import tempfile
from build_data import ROOT, SOURCE_NAME, build, read_json, atomic_bytes, encode

WORKFLOW='.github/workflows/fault-snapshot-v3.yml'
SOURCES=[SOURCE_NAME,'data/daily.xlsx']

def git(*args):
    return subprocess.check_output(['git',*args],cwd=ROOT)

def main():
    statepath=ROOT/'data/v3/git-sync.json'
    state=read_json(statepath,{'revisions':[]})
    done=set(state['revisions'])
    # Latest branch state, not the potentially stale push-event checkout.
    head=git('rev-parse','HEAD').decode().strip()
    installs=git('log','--reverse','--diff-filter=A','--format=%H','HEAD','--',WORKFLOW).decode().splitlines()
    start=installs[0] if installs else head
    changes=git('log','--reverse','--format=%H',f'{start}..{head}','--',*SOURCES).decode().splitlines()
    revisions=list(dict.fromkeys([start]+changes))
    processed=0
    for rev in revisions:
        if rev in done:continue
        candidate=None
        for path in SOURCES:
            try:
                raw=git('show',f'{rev}:{path}')
                candidate=raw;break
            except subprocess.CalledProcessError:continue
        if candidate is None:
            # No workbook at installation is fine; first upload is processed later.
            done.add(rev);continue
        stamp=git('show','-s','--format=%cI',rev).decode().strip()
        with tempfile.TemporaryDirectory() as temp:
            workbook=Path(temp)/'input.xlsx';workbook.write_bytes(candidate)
            result=build(ROOT,workbook,rev,stamp)
            print(json.dumps(result,ensure_ascii=False))
        done.add(rev);processed+=1
    # A workflow/code-only rerun must also regenerate the current snapshot with new logic.
    if not processed:
        from build_data import choose_source
        try:source=choose_source(ROOT)
        except ValueError:
            raise SystemExit('먼저 원본 엑셀을 기존 이름으로 올려주세요.')
        result=build(ROOT,source,head)
        print(json.dumps(result,ensure_ascii=False))
    state['revisions']=sorted(done)
    atomic_bytes(statepath,encode(state))

if __name__=='__main__':main()
