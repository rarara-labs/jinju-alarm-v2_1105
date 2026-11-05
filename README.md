# GN Fault Insight V2 · 모바일 고장 대시보드

## 먼저 확인하세요 (중요)
- **외부 서비스 반출 승인이 있는 데이터만 사용**하세요. Private GitHub 저장소와 비공개 Streamlit 앱도 사내 보안정책을 대신하지 않습니다.
- **기존 저장소/Streamlit 앱은 수정하지 않습니다.** 새 시험 저장소를 만드세요.
- 과거 알람을 완전히 보존하려면 **기존 저장소 `alarm_history.json`을 내려받아 새 저장소 `data/previous_alarm_history.json` 이름으로 올려야 합니다.** 이 파일은 V2 새 이력과 자동 병합되므로 기존 이력도 새 이력도 잃지 않습니다. 기존 파일이 없으면 엑셀에 남은 이력부터 새로 축적되며 과거 기록 전체 복원을 보장하지 못합니다.

## GitHub에 올리기
1. GitHub에서 새 **Private** 저장소 `jinju-alarm-v2`를 생성합니다.
2. ZIP 압축을 **풀고**, 폴더 안의 파일/폴더를 저장소 루트에 업로드합니다. `.github/workflows/build.yml`과 `.streamlit/config.toml` 같은 숨김 폴더도 포함해야 합니다.
3. **기존 이력 이전:** 기존 저장소의 `alarm_history.json`을 내려받아 **`previous_alarm_history.json`**으로 이름 변경 후 새 저장소 `data/` 폴더에 업로드하면 다음 자동 실행 시 병합됩니다. `data/alarm_history.json`은 덮어쓰지 마세요.
4. GitHub의 Actions 탭에서 `Build daily alarm snapshot`이 **초록색 성공**인지 확인하세요. 첫 업로드 시 자동으로 실행됩니다. 실패하면 Actions 메뉴에서 오류 내용을 확인하세요.
5. GitHub Actions가 저장 결과 `data/processed.json`, `data/alarm_history.json`을 커밋했는지 확인하세요. 완료가 확인되지 않으면 엑셀을 다시 교체하거나 기존 이력을 삭제하지 마세요.
6. Streamlit Community Cloud (https://share.streamlit.io/) → **Create app** → 새 저장소 → Branch `main` → Main file `app.py` → **Deploy**.

## 매일 갱신하기
- GitHub 저장소의 `data/daily.xlsx`를 **같은 이름으로 교체**해 Commit 합니다.
- 커밋 시 GitHub Actions가 모든 관련 RAW에서 금일 알람명을 추출하고 기존 `data/alarm_history.json`에 신규 값을 병합하여 자동 보존합니다.
- **날짜별로 저장된 알람은 기존 값이 있으면 덮어쓰지 않습니다.** 업무상 수정이 필요한 경우 별도 검토가 필요합니다.
- GitHub Actions 성공 후 대시보드 상단 **기준일**을 확인하고 필요하면 **최신 자료 확인** 버튼을 누릅니다.
- 실패했을 때 이전 스냅샷이 유지되므로, 아침 업로드 후 꼭 처리 성공 여부를 확인해 주세요.

## GitHub Actions 실패 시
- Repo Settings → Actions → General → **Workflow permissions**에서 `Read and write permissions`를 사용할 수 있는지 확인합니다. Workflow에도 `contents: write`가 명시되어 있습니다.
- Actions 실행 권한이 막혀 있는 조직 계정일 수 있습니다. 이 경우 관리자 정책을 확인하세요.
- Actions → Build daily alarm snapshot → Run workflow 로 수동 재실행할 수 있습니다.

## 성능 설계
- 휴대폰은 Excel(수 MB)을 읽지 않고 처리된 `data/processed.json`만 열어 카드/표를 출력합니다.
- Excel 분석, RAW 매칭, 알람 이력 파일 쓰기는 GitHub Actions가 업로드 시 1회 수행합니다.
- 앱의 콜드 스타트(절전 해제) 시간은 Streamlit Cloud 환경에 따라 별도로 발생할 수 있습니다.

## 주의: 데이터 분류 및 정확도
- 알람 매칭은 기존 앱의 우선순위(DU/RF → DU → OOS 값)를 가능한 범위에서 보존했으나, **PORT 매핑은 원본 매핑표 구조에 따라 일부 확인 필요가 남을 수 있습니다.** 화면의 `알람 근거`를 확인하세요.
- 날짜 기준은 OOS 시트의 `본부` 열입니다. 동일 PORT라도 날짜별 알람은 구분해 저장합니다.
- `금일`은 실제 오늘이 아니라 OOS 자료에 존재하는 **최신 기준일**입니다.
- 기존 앱의 14개 시트 중 모바일에 필요한 핵심 시트와 이력을 가공합니다. V1 기능의 완전한 1:1 교체는 아니며 **시범용 V2**입니다.
- 동일한 날짜에 엑셀을 여러 번 교체한 경우, 날짜별 알람은 먼저 저장된 값을 보존합니다. 필요시 별도 정정 기능이 있어야 합니다.
