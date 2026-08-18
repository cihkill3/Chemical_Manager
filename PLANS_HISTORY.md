# 📚 Chemical Manager 개발 계획서 누적 기록 (PLANS HISTORY)

본 문서는 Antigravity AI 코딩 어시스턴트가 수행하는 모든 작업의 사용자 지시 내용과 구현 계획서(Implementation Plan)를 **지속적으로 누적 기록(Append Only)**하는 공식 기록 보관소입니다.

---

## 📜 [2026-08-13 #1] OneDrive 호환성 낙관적 동시성 제어 및 3-Way 충돌 해결 계획서

### 1. 사용자 지시 사항 요약
- 원드라이브/공동편집 환경에서 파일이 열려 있거나 타인이 동시에 편집 중일 때 충돌 없이 작동하도록 조치.
- 동기화 및 DB 업데이트 시 덮어쓰기 직전 타인 편집 여부 및 최신 수정 내역을 검사하고, 편집 중이면 대기 재시도.
- 타인이 새로 작성한 내용과 내 변경사항이 겹치지 않으면 3-Way Merge, 겹치면 취소 후 최신 파일로 재동기화.

### 2. 세부 구현 계획 (Plan Details)
- **스냅샷 타임스탬프 저장**: 파일 오픈 시점의 `mtime`을 기록하고 저장 직전 재검증.
- **Pre-Save Lock Check**: 저장 직전 `is_file_locked` 여부를 재확인하고, 잠겨 있는 경우 백오프 대기.
- **3-Way Merge & Rollback**: 타인의 변경 Diff와 내 작업 셀이 겹치지 않으면 병합 저장, 겹치면 전면 취소(Rollback) 후 재시작.

---

## 📜 [2026-08-13 #2] 조건부 백업, 1분 주기 대기, 동기화/업데이트 중지 버튼 구현 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
1. **조건부 백업**: 수정할 내용이 없으면 백업 파일을 생성하지 않음. 일단 파일은 읽어둔 뒤 나중에 실제 동기화/업데이트할 내용이 발생하면 저장 직전에 백업을 수행하여 불필요한 백업 남발 차단.
2. **1분 대기 및 자동동기화 연동**: 파일이 편집 중일 때 2초 대기 대신 1분 간격으로 대기. 자동 동기화 설정 시 다음 자동 동기화 전까지만 1분 단위로 반복(실패 시 넘김), 미설정 시 중지할 때까지 1분 단위 무한 대기.
3. **작업 중지 버튼**: 수동 동기화 및 DB 업데이트 진행 중일 때 버튼을 활성화("🛑 동기화 중단" / "🛑 DB 업데이트 중단")하여 클릭 시 작업을 즉시 취소/정지.
4. **누적 계획서 및 절대 원칙 고정**: 본 계획서와 절대 원칙(`ABSOLUTE_PRINCIPLES.md`)을 지속 누적(Append Only) 기록하고 코딩 시 100% 참조.

### 2. 세부 구현 계획 (Plan Details)

#### [Component 1: Sync Engine & DB Manager (core/sync_engine.py)]
- **조건부 백업 수칙 적용**:
  - `run_sync()` 초기 타겟 오픈 단계의 백업 실행 코드를 제거.
  - 데이터 비교 및 크롤링 검토 후, `cnt_new > 0` 또는 `cnt_upd > 0`인 경우에만 `self.target_wb.Save()` 직전에 백업(`ChemicalList_backup_...xlsx`) 수행.
- **1분 대기 및 자동동기화 연동 잠금 처리**:
  - `check_and_wait_lock()` 함수 수정: 대기 주기를 `time.sleep(60)` (1분)으로 변경.
  - 자동 동기화 켜짐(`auto_sync` True): `max_attempts = max(1, interval_minutes - 1)` 동안 1분 단위 재시도 후 그래도 잠겨 있으면 예외 없이 해당 주기 스킵 로그 후 종료.
  - 자동 동기화 꺼짐(`auto_sync` False): 취소 요청이 올 때까지 `while not stop_requested:` 1분 단위 무한 대기.

#### [Component 2: GUI Main Window (gui/main_window.py)]
- **수동 동기화 & DB 업데이트 중지 버튼 활성화**:
  - 수동 동기화 시작 시: `self.btn_sync.setText("🛑 동기화 중단")`, `self.btn_sync.setEnabled(True)`.
  - DB 수동 업데이트 시작 시: `self.btn_db_update.setText("🛑 DB 업데이트 중단")`, `self.btn_db_update.setEnabled(True)`.
  - 중단 버튼 클릭 시:
    - 작업 스레드(`SyncWorker` / `DbUpdateWorker`)에 `stop()` 시그널 전송.
---

## 📜 [2026-08-13 #3] 타겟 파일 오픈 시 2초 대기 버그 수정 및 초기 백업 완전 제거 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
1. **타겟 파일 오픈 시 2초 대기 버그 수정**: `sync_engine.py` 초기 타겟 오픈 단계(L.143)에서 `retry_delay=2` 및 `max_retries=5`로 잘못 호출되던 부분을 수정하여, 타겟 오픈 시에도 1분 단위 대기(`retry_delay=60`) 및 자동동기화 주기 연동(`max_retries`)이 완전 적용되도록 보장.
2. **초기 백업 완전 제거**: `sync_engine.py` L.127~L.136에 잔존해 있던 초기 오픈 시점의 무조건 백업 코드를 완전 삭제하여, 변경사항 0건 시 백업 파일 생성을 100% 차단.

### 2. 세부 구현 계획 (Plan Details)
- **`sync_engine.py` L.127~136 삭제**: 초기 타겟 열기 시점의 백업 실행 코드를 완전히 제거.
- **`sync_engine.py` L.143 수정**: `check_and_wait_lock` 호출 시 `retry_delay=60`, `is_auto_sync=is_auto_sync`, `max_retries=max_retries` 인자 전달.

---

## 📜 [2026-08-13 #4] DB 업데이트 중단 시 저장 차단 및 변경사항 0건 시 엑셀 오픈/저장 방지 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
1. **DB 변경사항 없을 때 저장 방지**: 크롤링 또는 업데이트 대상 데이터가 0건이거나 기존 DB와 차이가 없을 경우 `ChemicalList.xlsx` 파일 오픈 및 저장을 시도하지 않고 즉시 종료.
2. **중단 요청 시 DB 저장 완전 차단**: 중단 버튼 클릭 시 `is_stopped` 플래그를 확인하여 즉시 `success=False`로 반환하고, 후속 `non_conflicts` 처리 및 `save_db_records_win32com` 파일 저장 루틴 진입을 100% 차단.

### 2. 세부 구현 계획 (Plan Details)
- **`DbUpdateWorker.run()` 개편 (`gui/main_window.py`)**:
  - 크롤링 루프 종료 후 `if getattr(self, "is_stopped", False):` 검사 추가. 참일 경우 `success=False, error="사용자에 의해 DB 수동 업데이트가 중단되었습니다."` 반환 후 전면 종료.
  - `crawled_results_batch`가 비어있거나 `updated == 0`인 경우 `non_conflicts`를 빈 리스트로 반환 후 안내 로그 출력.
- **`on_db_update_finished()` 개편 (`gui/main_window.py`)**:
  - `records_to_save`가 빈 리스트일 경우 `save_db_records_win32com`을 호출하지 않고 `"업데이트할 새로운 DB 정보가 없습니다. (파일 미변경)"` 알림 표시 후 리턴.

---

## 📜 [2026-08-13 #5] DB 재크롤링 데이터의 실질 변경사항(Real Change) 감지 및 동일 데이터 저장 차단 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
- 이미 DB에 존재하는 시약 정보를 재크롤링했을 때, 크롤링된 결과가 기존 DB 레코드와 100% 동일하다면 변경사항(`non_conflicts`)으로 간주하지 않음.
- 동일 내용 재반영으로 인한 불필요한 엑셀 파일 저장 및 `"총 N개 반영"` 오인 메시지 발생을 차단.

### 2. 세부 구현 계획 (Plan Details)
- **`DbUpdateWorker.run()` 비충돌 처리 개편 (`gui/main_window.py`)**:
  - `conflicting_fields`가 없을 때, 기존 DB 값(`existing`)과 크롤링 값(`db_result`) 간 실질적인 값 변경(`has_real_change`)이 발생했는지 검사.
  - 새로 채워지거나 값이 변경된 필드가 단 하나도 없다면 `non_conflicts`에 추가하지 않고 스킵.
  - 실질적인 변경이 존재하는 경우에만 `non_conflicts`에 추가하여, 최종 `records_to_save`가 0건일 경우 엑셀 저장을 완전히 차단.

---

## 📜 [2026-08-13 #6] DB 수동 업데이트 시작 시 및 DB 저장 시 1분 대기(check_and_wait_lock) 적용 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
1. **DB 수동 업데이트 시작 시 잠금 감지 미흡**: DB 업데이트 시작 단계에서 `is_file_locked` 감지 후 1초 대기하고 무조건 크롤링으로 넘어가던 오류 수정. 1분 단위 대기(`check_and_wait_lock`)로 전환하여 파일이 닫힐 때까지 대기.
2. **DB 저장 시 예외 직행 버그 수정**: 크롤링 완료 후 DB 저장을 수행하는 `save_db_records_win32com`에서 파일이 읽기 전용(`ReadOnly`)일 때 예외를 즉시 던지던 버그 수정. `check_and_wait_lock`을 통해 1분 단위로 대기 후 잠금이 해제되면 안전하게 저장.

### 2. 세부 구현 계획 (Plan Details)
- **`DbUpdateWorker.run()` 초기 단계 개편 (`gui/main_window.py`)**:
  - `is_file_locked` 감지 시 1초 대기 대신 `check_and_wait_lock(target_path, log_fn=self.progress.emit, retry_delay=60, ...)`을 호출하여 파일 잠금 해제 1분 단위 대기.
- **`save_db_records_win32com` 개편 (`core/sync_engine.py`)**:
  - `wb.ReadOnly` 감지 시 예외를 바로 던지지 않고 `check_and_wait_lock`으로 1분 단위 대기 수행 후 파일이 닫히면 Excel 재오픈 및 저장 완료.

---

## 📜 [2026-08-13 #7] ChemicalList.xlsx 내 Guide 시트 미려한 서식 적용 및 자동 동기화 연동 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
- `excel_user_guide.md` 파일의 내용을 `ChemicalList.xlsx` 엑셀 파일 내 `Guide` 시트에 미려하고 보기 편한 시각적 디자인(헤더 배너, 표 서식, 경고 상자, 세련된 색상)으로 배치.

### 2. 세부 구현 계획 (Plan Details)
- **Visual Design Spec**:
  - **메인 타이틀 배너 (A1:D2 Merged)**: Deep Royal Navy Blue (`#1E3A8A`) 배경, Bold 15pt 흰색 텍스트, 중앙 정렬.
  - **섹션 배너**: Sky Blue (`#0284C7`) 배경, Bold 12pt 흰색 텍스트.
  - **서브 헤더**: Light Ice Blue (`#EFF6FF`) 배경, Bold 11pt Deep Blue 텍스트.
  - **표 헤더**: Blue Tint (`#DBEAFE`) 배경, Thin Gray Border (`#CBD5E1`).
  - **경고 카드 Box**: Soft Pastel Red (`#FEE2E2`) 배경, Dark Red (`#991B1B`) 텍스트, Left Border.
  - **데이터 행**: Alternating Soft Gray (`#F8FAFC`) / White (`#FFFFFF`), Consolas font for catalog/CAS examples.
- **Win32COM & Openpyxl 지원 스크립트 작성**: `scratch/update_guide_sheet.py` 작성 및 `ChemicalList.xlsx` 반영 완료.

---

## 📜 [2026-08-13 #8] Guide 시트 및 excel_user_guide.md 4가지 수칙 개편 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
1. **Index 시트 안내 추가**: 시트 구성 설명에 `Index` 시트 (시약 색인판: 시약명, CAS No., 위치 빠른 검색) 항목 추가.
2. **Status 직접 기입 항목 삭제**: `ChemicalList` 직접 기입 항목에서 `Status` 삭제 (`Quantity`와 `Used` 수량 비교를 통해 `O`/`X` 자동 판정).
3. **Orderbook 규칙에서 CAS 번호 삭제**: Orderbook 엑셀에서 CAS 번호 필드를 읽지 않으므로 주요 항목 기입 규칙 표에서 삭제.
4. **Ctrl+S 저장 안내 문구 삭제**: 동기화가 실시간 저장 감지가 아닌 주기에 맞춰 주기적으로 작동하므로 저장 당부 문구 삭제.

### 2. 세부 구현 계획 (Plan Details)
- **`excel_user_guide.md` 및 `scratch/update_guide_sheet.py` 개편**:
  - `Index` 시트 역할 설명 추가.
  - `Status` 항목 제거 및 `Used` 설명 보완 (`Quantity` 비교 자동 판단).
  - Orderbook 기입 표에서 `CAS 번호` 행 삭제.
  - Orderbook 주의사항에서 `Ctrl + S 저장` 안내 삭제.

---

## 📜 [2026-08-13 #9] 동일 시약 재크롤링 시 메타데이터(Revision Date)로 인한 오검출 저장을 차단하는 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
- 이미 DB에 존재하는 동일 시약을 재크롤링할 때, 실제 시약 정보에는 변경이 없는데도 매번 저장되고 `"총 1개 반영"` 알림이 발생하던 버그 수정.
- 본 원칙을 `ABSOLUTE_PRINCIPLES.md`에 [원칙 9]로 등록하고 코드에 100% 반영.

### 2. 세부 구현 계획 (Plan Details)
- **`DbUpdateWorker.run()` 비충돌 처리 개편 (`gui/main_window.py`)**:
  - `has_real_change` 판단 시 단순 `Revision Date`나 시스템 메타데이터 변경은 거르고, 실질 화학 정보 필드(`fields_to_check`: Product Name, CAS No., Storage Temp., Sensitivity, Signal Word, Key Hazards 등)에 값 변경이 있는 경우에만 `has_real_change = True`로 설정.
  - 실질적인 정보 변경이 없을 경우 `non_conflicts`에 추가하지 않아 엑셀 파일 저장을 100% 차단.

---

## 📜 [2026-08-13 #10] SDS PDF 파일 6개월(180일) 유효 보존 및 재다운로드 차단 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
- DB 업데이트 및 크롤링 수행 시, SDS 링크를 포함한 모든 시약 데이터가 기존과 동일하고, 기존 SDS PDF 파일이 존재하며 해당 파일의 작성/수정 일자가 오늘 기준 6개월(180일) 이내인 경우 SDS 파일 재다운로드 및 재저장을 차단.
- 본 수칙을 `ABSOLUTE_PRINCIPLES.md`에 [원칙 10]으로 고정 및 코드 반영.

### 2. 세부 구현 계획 (Plan Details)
- **`DBManager.is_sds_fresh(sds_path, max_days=180)` 신설 (`core/db_manager.py`)**:
  - 지정한 local SDS 파일의 존재 여부, 파일 크기, 그리고 수정 시각(`mtime`) 기준 경과 일수(`age_days < 180`)를 신속 검사하는 정적 메서드 구현.
- **각 제조사별 Scraper 개편 (`scrapers/aldrich.py`, `tci.py`, `thermofisher.py`)**:
  - PDF 바이트 검증 완료 후, 기존 SDS PDF 파일이 존재하고 `is_sds_fresh`가 `True`인 경우 `open(sds_path, 'wb')` 덮어쓰기를 생략하고 기존 파일 경로 유지.

---

## 📜 [2026-08-13 #11] is_new_target 변수미정의 NameError 수정 및 원본 오더북 무조건 ReadOnly 오픈 처리 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
- 수동 동기화 진행 시 `name 'is_new_target' is not defined` 예외 발생 버그 수정.
- 사용자의 원본 오더북 파일은 어떠한 경우에도 엑셀 쓰기 잠금이 발생하지 않도록 무조건 `ReadOnly=True` 읽기 전용 모드로만 오픈하는 수칙을 `ABSOLUTE_PRINCIPLES.md`에 [원칙 11]로 등록.

### 2. 세부 구현 계획 (Plan Details)
- **`is_new_target` 정의 추가 (`core/sync_engine.py`)**:
  - Target Setup 진입 단계에서 `is_new_target = not (target_path and os.path.exists(target_path))`를 명시적으로 정의하여 NameError 오류 수정.
- **원본 파일 `ReadOnly=True` 강제 오픈 (`core/sync_engine.py`)**:
  - `self.excel.Workbooks.Open(src_path, ReadOnly=True, UpdateLinks=False)`로 설정하여 원본 파일 수정/잠금 방지.

---

## 📜 [2026-08-13 #12] 파일 잠금 대기 중 중단 클릭 시 0.1초 즉시 중단 처리 계획서

### 1. 사용자 지시 사항 요약 (User Directive Summary)
- 수동 동기화 및 DB 업데이트 중단 클릭 시 1분 대기 루프에서 `check_stop_fn`이 누락되거나 1초 단위 블록킹으로 인해 중단 반응이 지연되던 버그 수정.
- 사용자가 중단 버튼을 클릭하는 즉시 0.1초 이내에 대기를 파괴하고 작업을 즉시 반납/종료하도록 개편.
- 본 수칙을 `ABSOLUTE_PRINCIPLES.md`에 [원칙 12]로 고정 및 코드 반영.

### 2. 세부 구현 계획 (Plan Details)
- **`check_and_wait_lock()` 0.1초 슬립 개편 (`core/sync_engine.py`)**:
  - 대기 루프 내에서 `time.sleep(1)` 대신 `for _ in range(retry_delay * 10): time.sleep(0.1)`로 변경하여 클릭 시 100ms 이내 반응 및 반환.
- **`check_stop_fn` 바인딩 파이프라인 전달 (`core/sync_engine.py` & `gui/main_window.py`)**:
  - `SyncEngine.__init__`, `SyncWorker`, `save_db_records_win32com`에 `check_stop_fn` 파라미터를 100% 바인딩하여 `is_stopped` 클릭 시 대기 루프를 0.1초 만에 즉각 이탈.

---

## 📜 [2026-08-13 #13] OneDrive 공동편집 3-Way Commit 및 20,000행 동적 확장 구현 계획서

### 1. 사용자 지시 사항 요약
- `ABSOLUTE_PRINCIPLES.md`와 기존 누적 계획을 기준으로 OneDrive에 저장된 `ChemicalList.xlsx`를 여러 사용자가 공동편집하는 상황을 전제로 구현.
- 분석 내용을 두 누적 Markdown 문서에 Append Only 방식으로 반영.
- 전체 행이 1,000행을 넘을 가능성이 높으므로 서식, 유효성 검사 및 수식 준비 범위를 넉넉하게 확장.

### 2. 구현 내용

#### [Component 1: 낙관적 동시성 모듈 (`core/concurrency_manager.py`)]
- 공유본의 `mtime_ns`, 파일 크기, SHA-256을 묶은 로컬 파일 지문 도입.
- ChemicalList는 `Order No.`, DB는 정규화된 `Manufacturer|Catalog No.` 기준으로 행 스냅샷 생성.
- `base / ours / theirs` 셀 단위 3-Way 비교 및 비충돌 패치 생성.
- 동일 셀의 서로 다른 변경, 공동편집자 행 삭제, 헤더 구조 변경을 직접 충돌로 차단.
- 프로세스 내부 workbook 쓰기 작업을 하나로 제한하는 공통 잠금 추가.

#### [Component 2: 동기화 엔진 작업 복사본 및 Commit (`core/sync_engine.py`)]
- 기존 공유본을 시스템 임시 폴더에 복사하고, 주문서 비교·크롤링·SDS·수식 계산은 작업 복사본에서만 수행.
- 저장 직전에 최신 공유본을 다시 스냅샷하여 3-Way merge하고 비충돌 셀만 최신본에 적용.
- 직접 충돌 시 공유본 미저장 상태로 최신본에서 최대 2회 전체 재시작.
- 백업은 실제 셀 패치가 존재할 때 최신 공유본을 기준으로 저장 직전에 1회 생성.
- 중복 `cleanup()` 제거, 정상 `Workbook.Close()`/`Excel.Quit()`만 사용하고 자기 Excel PID `taskkill` 완전 제거.
- 원본 주문서는 계속 `ReadOnly=True`로만 사용하고 저장 코드를 제거.

#### [Component 3: DB 수동 업데이트 공동편집 보호 (`gui/main_window.py`, `core/sync_engine.py`)]
- DB 크롤링부터 충돌 선택 및 최종 저장까지 프로세스 내부 단일 쓰기 작업 잠금 유지.
- DB 저장도 공유본이 아닌 작업 복사본에서 먼저 작성한 뒤 최신 공유본과 3-Way merge.
- 직접 충돌 시 기존 크롤링 결과로 강제 저장하지 않고 최신 DB를 다시 읽어 최대 2회 전체 DB 업데이트 재실행.
- 자동 동기화 판단 키를 `sync_interval_minutes`로 통일.

#### [Component 4: 대용량 행 범위]
- 기본 범위를 1,000행에서 최소 20,000행으로 확대.
- 실제 마지막 행이 15,000행을 넘으면 `마지막 행 + 5,000행`까지 자동 확장.
- 최신 공유본 commit 단계에서는 기존 수식을 광범위하게 재입력하지 않고 작업 복사본에서 검증된 셀 패치만 적용하여 사용자 셀 값을 보존.

#### [Component 5: 검증]
- 공개 함수 import 무결성 검사.
- 비중첩 공동편집 보존, 동일 셀 충돌, 동일 결과 저장 생략, 행 삭제 충돌, 20,000행 동적 범위에 대한 자동 테스트 추가.
- 실제 OneDrive 클라우드 반영 지연과 서버 리비전은 로컬 파일 지문만으로 완전한 원자성을 보장할 수 없으므로, 실제 배포 검증에서는 Microsoft 365 최신 Excel/OneDrive.exe 환경과 두 계정을 이용한 공동편집 시나리오를 필수 수행.

### 3. 구현 후 보완 및 검증 결과
- DB 수동 업데이트는 크롤링 시작 시점의 workbook 스냅샷을 최종 저장 함수까지 전달한다. 따라서 크롤링 도중 공동편집자가 같은 DB 필드를 변경하면 저장 직전 직접 충돌로 판정되며 해당 값을 덮어쓰지 않는다.
- DB 최종 저장은 GUI 스레드가 아닌 `DbSaveWorker`에서 실행한다. 최신본 검증·잠금 대기·20,000행 관리 범위 갱신 중에도 화면과 중단 버튼이 응답하며, 저장 직전까지 중단 요청을 재확인한다.
- 데이터 변경이 0건이어도 기존 유효성 검사 또는 조건부 서식 범위가 목표 행보다 작을 때만 이를 실제 변경으로 판정하여 20,000행 범위를 1회 확장한다. 이미 충분한 범위이면 저장과 백업을 모두 생략한다.
- 조건부 서식은 매번 중복 추가하지 않고 프로그램 관리 범위를 한 번 재구성하여 규칙 누적을 방지한다.
- Python 컴파일 및 공개 함수 import 검사를 통과했다.
- 순수 동시성 자동 테스트 8건을 통과했다: 비중첩 변경 보존, 동일 셀 충돌, 동일 결과 저장 생략, 공동편집 행 삭제 충돌, 헤더 변경 충돌, 신규 DB 시트 병합, 20,000행 동적 계산, 기존 metadata 범위 탐지.
- 현재 개발 PC의 Excel COM 자동화 세션은 별도 임시 workbook 통합 시험에서 비정상적으로 장시간 대기하여 완료하지 못했다. 시험 중 생성된 숨은 Excel 프로세스와 임시 로그는 정리했으며 운영 `ChemicalList.xlsx`는 시험에 사용하거나 수정하지 않았다. 실제 OneDrive 두 계정 통합 시험은 배포 전 필수 잔여 검증으로 유지한다.

---

## 📜 [2026-08-13 #14] 크롤링 선행 최적화 1·2·3·5 구현 기록

### 1. 사용자 지시 사항
- 고정 대기를 조건부 대기로 변경.
- 최신 SDS를 다운로드 전에 검사하여 네트워크 요청 생략.
- 정규화된 제조사·카탈로그 번호 기준으로 중복 크롤링 제거.
- 크롤링 결과 수집과 Excel 기록 단계를 분리.
- 모든 수정 사항을 Markdown 문서에 누적 기록.

### 2. 구현 내용
- `BaseScraper.wait_for_page()`를 추가하여 `document.readyState`, 제목, 선택 요소를 0.1초 간격으로 확인하고 준비 즉시 반환하도록 구현.
- Abcam 5초, Aldrich 3초, TCI 3초/2초, Thermo Fisher 2초/3초 고정 대기를 조건부 대기로 교체.
- `BaseScraper.find_fresh_sds()`를 추가하고 제조사·카탈로그가 일치하는 180일 이내 PDF를 네트워크 요청 전에 검색.
- Aldrich, TCI, Thermo Fisher에서 최신 로컬 SDS 발견 시 SDS 탐색 API와 PDF 다운로드를 생략하고 기존 절대경로를 반환.
- `DBManager.crawl_key()`를 추가하여 제조사 별칭, 대소문자, Excel 숫자형 카탈로그의 `.0` 차이를 동일 키로 정규화.
- DB 수동 업데이트의 작업 사전을 정규화된 키로 구성하여 ChemicalList와 DB에 중복된 제품을 한 번만 크롤링.
- 전체 동기화는 기존 DB 키를 먼저 수집하고, 누락된 고유 제품의 크롤링 결과를 `crawl_results`에 선반영한 뒤 Excel DB 셀 기록 루프를 실행하도록 분리.
- 전체 동기화 스크래퍼에도 `check_stop_fn`을 전달하여 조건부 대기 중 중단 요청을 처리.

### 3. 검증 결과
- 가상환경 Python으로 전체 `compileall` 통과.
- 기존 공동편집 테스트 8건과 신규 최적화 테스트 3건, 총 11건 통과.
- 신규 테스트는 정규화 중복 키, 조건부 대기의 조기 반환, 최신 SDS 선검사를 검증.
- 실제 제조사 사이트의 응답 시간과 차단 정책은 네트워크 통합 시험이 필요하며, 이번 단계에서는 제조사 내부 병렬 크롤링을 도입하지 않음.

---

## 📜 [2026-08-13 #15] 전체 코드 재점검 후 운영·배포·크롤링 구조 보완 기록

### 1. 사용자 지시 사항
- 전체 코드 재검토에서 확인된 수정 후보를 모두 반영.
- 운영 데이터와 백업 파일은 삭제하지 않고 코드·설정·빌드·테스트·문서만 수정.
- 모든 구현 내용을 Markdown에 누적 기록.

### 2. 구현 내용
- 실제 import에 맞춰 `PyQt6`, `pywin32`, `seleniumbase`, `pymupdf`를 런타임 의존성에 추가하고 개발용 `pytest`, `pyinstaller`는 별도 파일로 분리.
- `run.bat`의 프로젝트 루트 가상환경 경로를 수정하고 PyInstaller가 SeleniumBase 동적 모듈·데이터·바이너리를 수집하도록 spec 보강.
- 제조사 별칭 정규화, 지원 여부와 스크래퍼 선택을 `scrapers/registry.py`로 통합.
- 특정 제품 DB 업데이트 식별자를 카탈로그 단독 값에서 제조사+카탈로그 복합 키로 변경.
- SDS 검색은 DB의 기존 `SDS_Local_Path`를 먼저 검사하고, 폴더 PDF 목록은 디렉터리 변경 시 한 번만 다시 만드는 인덱스로 최적화.
- DB 대화상자 임시 복사 파일을 고유 이름으로 생성하고 정리 실패를 경고로 기록.
- 전체 동기화의 제품 후보 계산과 모든 웹 크롤링을 헤더·행·DB 셀 쓰기보다 먼저 수행하도록 이동. 네트워크 완료 후 단일 Excel 기록 단계 시작.
- COA 다운로드 경로의 잔여 고정 대기를 DOM 준비 조건부 대기로 교체.
- 자동 동기화 중 수동 쓰기 버튼을 잠그고, 종료 시 모든 worker에 중단 신호를 전달한 뒤 제한 시간 동안 대기하도록 구현.
- 설정 저장 실패를 예외로 전달하여 GUI 오류로 표시하고, 기본 설정 반환은 중첩 객체까지 복사하도록 수정.
- 루트에서 테스트를 실행해도 import가 성립하도록 테스트 경로 설정을 보강하고 임시·로그·빌드 결과 ignore 규칙 추가.

### 3. 검증 결과
- 전체 Python 소스와 PyInstaller spec 정적 컴파일 통과.
- 루트 디렉터리 기준 unittest 13건 전부 통과.
- `pip check` 결과 설치된 패키지 의존성 충돌 없음.
- 프로젝트 가상환경에 누락되어 있던 `PyQt6`, `pywin32`를 설치하고 GUI·Excel·동시성·스크래퍼 공개 모듈 import 검사를 통과함.
- 루트 기준 `pytest` 실행 결과도 13건 전부 통과.
- 제조사 실제 웹사이트, Excel COM, 두 OneDrive 계정 공동편집, PyInstaller 최종 실행 파일은 운영 환경 통합 시험이 별도로 필요함.

---

## 📜 [2026-08-14 #16] Abcam SDS 및 관리 범위·DB 로그·갱신일 보완 기록

### 1. 사용자 지시 사항
- Abcam `ab92536` 제품 페이지의 SDS 링크를 찾아 PDF가 자동 다운로드되도록 수정.
- 실제 데이터가 약 20행일 때 관리 범위를 20,000행까지 확장하지 않도록 수정.
- 전체 동기화 로그에 DB 정보 보완 대상의 제조사와 카탈로그 번호를 표시.
- DB `Revision Date`에서 시간 부분을 제거하고 날짜만 저장·표시.

### 2. 구현 및 검증
- Abcam 공식 `proxy-gateway.abcam.com/product` GraphQL API와 `X-Abcam-App-Id` 헤더를 사용해 한국 SDS 목록을 조회하고 영문 문서를 우선 다운로드하도록 구현.
- S3 서명 URL에서 PDF를 내려받아 PDF 시그니처를 검증한 뒤 표준 SDS 파일명으로 저장.
- 조건부 서식과 유효성 검사 범위를 실제 마지막 데이터 행까지만 관리하도록 변경.
- DB 보완 완료 로그를 `제조사 카탈로그번호` 목록과 함께 출력하도록 변경.
- 최종 3-way commit의 DB 서식 단계에서 모든 갱신일을 자정 날짜 값과 `yyyy-mm-dd` 표시 형식으로 정규화.
- 전체 자동 테스트 16건 통과. 실제 `ab92536` 크롤링에서 영문 KGHS SDS PDF 326,903바이트 다운로드 성공.

## 17. COA/CoC 통합 및 동기화 책임 단순 분리 (2026-08-14)

### 요청과 구현

- 오더북 `Lot No.`를 기준으로 제조사 COA를 내려받아 ChemicalList의 이름 기반 열에 기록하도록 연결.
- `core/coa_manager.py`에 지원 제조사 판정, lot 정규화, cache 확인, 요청 중복 제거, 기존 downloader 호출, 만료일 추출, Abcam CoC 표시를 분리.
- `core/excel_manager.py`에 실제 헤더 탐색과 마지막 사용 열 뒤의 안전한 신규 헤더 예약/쓰기를 분리.
- Abcam CoC는 COA 열에 기록하되 `_CoC.pdf` 파일명을 유지하고 두 링크/경로 셀을 주황색으로 표시. 메모에는 CoC 대체 문서임과 Datasheet 링크/로컬 경로를 기록.
- CoC 주황색과 메모는 작업 복사본뿐 아니라 3-way 병합 직후 최신 공유본에도 다시 적용.
- 설정에 `Lot No.`, `Expiration Date`, `COA Link`, `COA Local Path`를 추가하고 기존 설정은 누락 항목을 로드 시 보완.

### 절대 원칙과 검증

- 원본 오더북 읽기 전용, OneDrive 외부 작업 복사본, key 기반 3-way 병합, 저장 직전 fingerprint 확인, 실제 변경 전 백업 순서를 유지.
- COA 네트워크 작업은 Excel 쓰기 전에 완료하며 실패 시 기존 정상 값을 지우지 않음.
- 새 헤더는 실제 마지막 헤더 뒤에만 추가하여 사용자 열과 충돌하지 않음.
- portable 및 앱 내 COA downloader는 수정하지 않았고 SHA-256으로 불변을 확인.
- 변경 전 19개, 변경 후 23개 pytest 통과. 수정 모듈 `py_compile` 통과.

## 18. Thermo Fisher A10436 오염 페이지 및 링크 교정 (2026-08-14)

### 원인

- `chemicals.thermofisher.kr`의 추측 suffix URL(`A10436.14`)이 제품 페이지 대신 CSS asset 내용을 HTTP 200으로 반환했지만, 기존 코드는 비어 있지 않은 `h1`만 보고 제품명으로 채택했다.
- fresh SDS가 있으면 실제 SDS URL 대신 Thermo Fisher 검색 결과 URL을 `SDS_Link`에 기록했다.
- 이미 DB에 존재하는 제품 키는 무조건 재크롤링 제외되어 오염된 행이 자동 교정되지 않았다.

### 수정

- 공식 글로벌 제품 URL `https://www.thermofisher.com/order/catalog/product/{catalog}`을 우선 사용.
- 제품명 길이, CSS marker 부재, page source의 catalog 번호 일치를 모두 검증한 뒤에만 제품 페이지로 인정.
- chemicals suffix URL은 엄격한 검증을 통과할 때만 fallback으로 사용.
- fresh SDS를 재사용해도 검증된 실제 PDF URL을 `SDS_Link`에 기록하고 검색 결과 URL은 저장하지 않음.
- CSS 오염 제품명 또는 잘못된 Thermo Fisher 검색 SDS 링크가 있는 기존 DB 행만 선택적으로 재크롤링하고 정상/수동 행에는 영향을 주지 않음.
- Revision Date의 시간 포함 문자열은 기존 날짜 정규화 경로로 `YYYY-MM-DD`만 보존.

### 검증

- 공식 페이지에서 A10436 제품명이 `Alexa Fluor™ 488 Hydrazide`, canonical URL이 `/order/catalog/product/A10436`임을 확인.
- 공식 `A10436_MTR-NALT_EN.pdf` URL이 HTTP 200 `application/pdf`임을 확인.
- CSS asset 거부, canonical 제품 승인, fresh SDS의 direct URL 유지, 기존 오염 행 재크롤링, 날짜-only 정규화 테스트 추가.
- 전체 pytest 29건 통과 및 수정 모듈 `py_compile` 통과.

## 19. 무변경 자동 동기화 반복 백업 방지 (2026-08-14)

### 원인

- Excel이 데이터 유효성 검사를 x14 확장 형식으로 저장한 파일에서 openpyxl은 해당 규칙을 읽지 못하고 빈 목록으로 반환한다.
- 기존 관리 범위 판정은 `validation_max < required_row OR conditional_max < required_row`였기 때문에 조건부 서식이 이미 최신이어도 validation 최대 행을 항상 0으로 판단했다.
- 그 결과 무변경 동기화마다 `force_metadata=True` 커밋, 백업, 저장이 반복됐다.

### 수정과 검증

- 조건부 서식과 유효성 검사는 같은 메서드에서 같은 관리 행 경계로 함께 적용되므로, openpyxl에서 안정적으로 읽을 수 있는 조건부 서식을 관리 범위 sentinel로 사용한다.
- 조건부 서식 범위가 현재 행까지 도달했으면 x14 validation이 보이지 않아도 추가 저장하지 않는다.
- 실제 ChemicalList는 데이터와 관리 조건부 서식 모두 60행까지 적용된 상태임을 읽기 전용으로 확인했다.
- validation 목록이 비어 있고 조건부 서식만 60행까지 있는 회귀 사례를 추가했다.

## 20. 제품별 네트워크 진행 로그 및 즉시 중단 (2026-08-14)

- 오더북의 `주문회사`와 제조사 매핑용 `회사`는 별개라는 사용자 운영 규칙에 따라 헤더/매핑은 변경하지 않음.
- DB 보완 네트워크 단계에서 `DB 정보 보완 중 (현재/전체): 제조사 카탈로그` 로그를 각 제품 시작 전에 출력.
- COA 단계에서도 `COA/CoC 다운로드 중 (현재/전체): 제조사 카탈로그 / Lot` 로그를 출력.
- 제품 크롤러의 차단형 HTTP 요청과 브라우저 async 요청을 daemon 작업으로 격리하고 호출 스레드가 0.1초마다 중단 상태를 확인하도록 변경.
- 기존 portable COA downloader는 수정하지 않고 `COAManager` 호출 계층에서 같은 0.1초 중단 polling을 적용.
- 중단 시 vendor timeout이 끝날 때까지 기다리지 않고 즉시 동기화 호출 스레드로 복귀하도록 회귀 테스트 추가.

## 21. 주문번호가 비어 있는 기존 파일의 행 병합 및 DB 수식 복구 (2026-08-14)

- ChemicalList의 신규 행 위치를 `Order No.` 열 하나로 판단하지 않는다. 기존 행의 주문번호가 비어 있을 수 있으므로 Order No., Product Name, Original Product Name, Manufacturer, Catalog No. 중 실제 마지막 행의 최댓값을 사용한다.
- 논리 키 기반 3-way 병합 후, 최신 공유본의 실제 행 번호와 헤더 이름을 기준으로 프로그램 관리 DB 조회 수식을 다시 생성한다. 주문번호가 없는 기존 행의 수식 변경은 일반 셀 패치만으로 전달할 수 없다.
- 사용자가 직접 입력한 CAS 값은 보존하며, 빈 값·실패 표시·기존 프로그램 수식인 경우에만 CAS 조회 수식을 재생성한다.
- 회귀 조건: Order No.의 마지막 행은 1이고 기존 시약 데이터는 60행까지일 때 신규 시약은 반드시 61행에 추가한다.

## 22. 전체 코드 재검토 후 운영 안전성 보완 (2026-08-14)

- 사용자의 명시적 판단에 따라 신규 행의 마지막 행 계산은 `Product Name` 열 기준을 유지했다.
- 조건부 서식·유효성 검사 관리 범위를 기본 3,000행으로 변경하고, 2,500행 초과 시 마지막 데이터 행보다 최소 1,000행 큰 값을 500행 단위로 올림하도록 변경했다.
- 자동 실행 옵션이라도 동기화 주기가 0분이면 시작 즉시 자동 동기화를 실행하지 않도록 수정했다.
- 주문번호가 없는 실제 기존 행에는 제조사·카탈로그·Lot·원본 제품명·중복 순번 기반 레거시 키를 부여하고, 수식만 있는 예약 행은 3-way 스냅샷에서 제외했다.
- DB 조회 수식과 Key 생성에서 고정 A:N 및 A/B/C열 의존을 제거하고 실제 헤더 열을 사용하는 `INDEX/MATCH` 방식으로 통일했다.
- COA 다운로드는 임시 격리 폴더에서 완료한 문서만 정식 폴더로 이동하며, 취소 시 브라우저 중단과 지연 정리를 수행하도록 보완했다.
- 숫자형 COA 만료일의 일/월 순서가 모호하면 잘못 추측하지 않고 빈 값으로 유지하도록 변경했다.
- 프로그램 종료 시 작업 스레드가 실제 종료되지 않았으면 애플리케이션 종료를 미루고 안전 종료를 반복 확인하도록 변경했다.
- 설정 JSON 파싱 실패를 기본 설정으로 조용히 대체하지 않고 사용자에게 오류로 전달하도록 변경했다.

## 23. Aldrich COA 선행 0 Lot 캐시 반복 다운로드 수정 (2026-08-14)

- I6256의 오더북 Lot `0000494416`이 ChemicalList에서 숫자 `494416`으로 저장되어 문자열 완전일치 캐시 판정이 매번 실패하는 원인을 확인했다.
- 기존 PDF `Aldrich_I6256_0000494416_COA.pdf`는 정상이고, 반복 실행마다 새 서명 URL만 달라져 비충돌 변경 1개 셀과 백업이 반복됐다.
- 숫자형 Lot 비교 키에서만 선행 0을 제거하고 원본 Lot 표시값과 파일명은 그대로 유지하도록 수정했다.
- 실제 I6256 PDF에 대해 Lot 동등성, PDF 유효성 및 결합 캐시 판정이 모두 True임을 검증했다.

## 24. 소분 시약명 보존 및 ChemicalList 기본 활성 시트 (2026-08-14)

- 주문번호 없는 행은 오더북 값 병합 대상이 아니지만 기존 DB 수식 재생성 과정에서 수동 Product Name이 바뀔 수 있는 경로를 확인했다.
- `Order No.`가 없고 Product Name이 리터럴 값이면 이름을 유지하며, 이미 DB 수식인 레거시 행만 수식을 갱신하도록 변경했다.
- Manufacturer, Catalog No., Lot No. 등 원 시약 추적 정보는 그대로 유지하고 Remarks의 Aliquot 표기는 사용자 값으로 보존한다.
- 작업 복사본 저장뿐 아니라 최신 공유본 3-way commit 직전에도 ChemicalList 시트를 활성화하도록 변경했다.
- 데이터 변경이 없어도 현재 활성 시트가 ChemicalList가 아니면 활성 시트 메타데이터를 한 번 저장하고, 이후에는 반복 저장하지 않도록 검사 로직을 추가했다.

---

## [2026-08-14 #25] 전체 프로그램 단일 실행파일 빌드 및 배포 검증 계획

### 사용자 지시
- Chemical Manager 전체 프로그램을 하나의 Windows 실행파일(EXE)로 빌드한다.
- 최근 분리된 동기화·동시편집·COA/CoC·제조사 크롤러 기능이 빌드에서 누락되지 않도록 검토한다.
- 빌드 전에 자동 테스트를 통과시키고, 빌드 후에는 실제 EXE 기동과 런타임 파일 동작을 검증한다.

### 구현 및 검증 계획
1. 기존 `Chemical_Manager.spec`의 one-file/windowed 설정을 유지한다.
2. SeleniumBase 데이터·바이너리·숨은 import와 제조사별 동적 scraper import를 포함한다.
3. 프로그램 안의 사용 가이드가 단일 EXE에서도 열리도록 `program_guide.md`를 번들 데이터에 포함한다.
4. 소스 전체 pytest와 compile 검사를 먼저 통과시킨다.
5. 이전 빌드 산출물과 분리된 임시 build/dist 경로에서 PyInstaller clean 빌드를 수행한다.
6. PyInstaller 경고 파일을 검토해 실제 런타임에 필요한 모듈 누락 여부를 판정한다.
7. 완성된 EXE를 임시 검증 폴더에서 기동하여 프로세스 생성, GUI 유지, 중복 실행 잠금, 정상 종료를 확인한다.
8. 원본 OneDrive Excel 파일은 빌드·기동 검증에서 열거나 저장하지 않는다.

---

## [2026-08-14 #26] 백업 3개월 보존·ChemicalList 선택·아이콘 통합 및 재빌드

### 사용자 지시
- 3개월 이상 지난 ChemicalList 백업 파일을 주기적으로 삭제한다.
- 프로그램 화면에서 사용할 ChemicalList.xlsx를 직접 선택할 수 있게 하고 전체 기능이 해당 파일을 사용하도록 검증한다.
- 제공된 ICO와 PNG를 프로그램 창, 트레이 및 EXE 아이콘으로 적용한다.
- 수정 완료 후 전체 회귀 테스트를 통과시키고 단일 EXE로 다시 빌드한다.

### 구현 및 안전 검증
1. 설정에 `target_file`을 추가하되 기존 설정은 오더북 폴더의 ChemicalList.xlsx로 호환한다.
2. 중앙 경로 결정 함수를 동기화, DB 업데이트, PDF 출력, SDS 병합에 공통 적용한다.
3. 선택 파일은 `.xlsx`, `ChemicalList`/`DB` 필수 시트, 원본과 다른 파일이라는 조건을 검사한다.
4. 매 동기화 시작 시 선택 파일 옆 `backup` 폴더의 정확한 백업 패턴만 확인하여 수정시각이 달력 기준 3개월 전 이하인 파일을 삭제한다.
5. 날짜 말일, 정확히 3개월, 최근 백업, 무관 파일 보존, 명시 대상 우선 및 레거시 경로 호환을 자동 테스트한다.
6. PNG는 Qt 창·트레이 런타임 자원으로, ICO는 PyInstaller EXE 자원으로 사용한다.
7. 전체 pytest·compile·Qt 오프스크린 UI·PyInstaller 빌드·EXE 기동을 순서대로 검증한다.

---

## [2026-08-15 #27] 프로그램 하위 백업 단일 경로 전환

### 사용자 지시
- 백업을 ChemicalList 파일의 하위 폴더가 아니라 로그와 같은 프로그램 하위 폴더에 저장한다.
- 주기 삭제를 포함하여 백업을 생성·조회하는 모든 로직을 검토하고 수정한다.

### 구현 및 검증
1. `get_app_root()/backup`을 반환하는 중앙 경로 함수를 `backup_manager.py`에 둔다.
2. 저장 직전 백업 생성도 중앙 함수로 이동하고 `sync_engine.py`의 대상 파일 기준 경로 조합을 제거한다.
3. 3개월 보존 정리는 대상 ChemicalList 경로를 받지 않고 프로그램 하위 backup만 검사한다.
4. 새 백업은 원본 파일의 수정시각을 상속하지 않도록 복사 직후 실제 생성 시각으로 수정시각을 설정한다.
5. ChemicalList가 프로그램과 다른 OneDrive 폴더에 있어도 프로그램 하위에만 백업이 생기고 OneDrive 쪽 backup은 생성되지 않는 회귀 테스트를 추가한다.
6. 전체 코드 검색으로 백업 경로 직접 조합이 중앙 모듈 밖에 남지 않았는지 확인하고 전체 테스트 및 컴파일을 수행한다.
