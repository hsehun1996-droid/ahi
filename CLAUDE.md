# 고속도로 포장유지보수 이력 관리 프로그램

## 프로그램 개요

고속도로 포장유지보수 공사 이력을 노선별로 관리하는 데스크탑 GUI 애플리케이션.  
Python + CustomTkinter(ctk) 기반, 라이트모드, 블루 액센트. 한국어 UI.

### 주요 기능
- **다중 노선 탭 관리**: 노선 추가/수정/삭제, 탭별 모식도 표시
- **모식도 시각화**: IC/JCT, 램프 포함. 구간별 공법 색상 표시, 돋보기(확대) 기능
- **이력 입력**: 공법(절삭 덧씌우기 등), 방향(상행/하행), 차로, 날짜, km 구간 (1m 단위 스냅)
- **하자기간 관리**: 카테고리별 하자기간 설정, 중복 시공 자동 감지(빗금 표시)
- **포장상태 관리**: DI지수, HPCI등급, AAR등급, RD등급, IRI등급
- **분석 기능**: 개량 우선순위 선정, 결함리스크 분석
- **사업계획**: 개량 우선순위 산정 → 구간 다중선택 → 사업계획 표 작성/수정(공법은 범례 공법만) → 확정 시 보수이력에 계획 엔트리(모식도에 '계획' 표기)로 반영
- **운영계획변경**: (보수필요 − 사업계획) 후보 산정 → 노선 선택(필수) → 노선명/사업명/목적/사업내용/단가 입력 → 한글(HWPX) 양식 자동 작성(zip+XML 직접 채움, 한글/COM 불필요)
- **구조물/IC 관리**: 교량·터널·IC·JCT·램프 이력
- **저장/내보내기**: CSV 저장·불러오기(자동 로드), PDF 내보내기, Excel 내보내기, 한글(HWP) 양식 내보내기
- **캐시**: SQLite DB(`highway_data.db`)로 빠른 재로드
- **입력 모드**: 본부/지사 모드 전환 (표시 항목 차이)

### 외부 라이브러리
- `customtkinter` — GUI 프레임워크
- `Pillow`, `reportlab` — PDF 내보내기 (선택)
- `openpyxl` — Excel 내보내기 (선택)
- `CTkMessagebox` — 다크모드 팝업 (선택)
- (운영계획변경 HWPX 내보내기는 표준 라이브러리 zip+XML만 사용 — 외부 라이브러리·한글·pywin32 불필요)

---

## 파일 구조 (리팩토링 후)

원본 단일 파일(`highway.py` 8,291줄)을 2026-04-10에 역할별로 분할 리팩토링함.

```
project1/
├── CLAUDE.md              ← 이 파일
├── highway.py             ← 진입점: MaintenanceApp.__init__ + if __name__ == "__main__"
├── constants.py           ← 모든 상수·설정값
├── utils.py               ← 유틸리티 함수 (비GUI)
├── canvas_utils.py        ← 캔버스 렌더링 전용 유틸
├── hwpx_export.py         ← 운영계획변경 한글(HWPX) 양식 자동 작성 (zip+XML 직접 채움)
├── templates/
│   └── operation_plan_template.hwpx ← 운영계획변경 한글 양식 템플릿(HWPX)
└── mixins/
    ├── __init__.py        ← 9개 Mixin 일괄 export
    ├── ui_mixin.py        ← UI 구성, 팝업 헬퍼, 대시보드, 공법/하자 설정
    ├── route_mixin.py     ← 노선 추가/수정/삭제, 탭 전환, 연도 필터
    ├── ic_mixin.py        ← IC·구조물·포장상태 데이터 다이얼로그
    ├── analysis_mixin.py  ← DI/HPCI/AAR/RD/IRI 분석, 우선순위, 결함리스크, 구간 이동
    ├── entry_mixin.py     ← 이력 입력 모드(본부/지사), 이력 추가
    ├── io_mixin.py        ← CSV·Excel·PS·PDF 저장/불러오기/내보내기, SQLite 캐시
    ├── window_mixin.py    ← 창 종료, 창 크기, 시작 시 자동 CSV 로드
    ├── canvas_mixin.py    ← 모식도 그리기, 돋보기/확대, 이력 상세 다이얼로그
    └── plan_mixin.py      ← 사업계획·운영계획변경 작성/확정, 한글 내보내기 연동
```

### 클래스 상속 구조

```python
# highway.py
class MaintenanceApp(
    UIMixin, RouteMixin, ICMixin, AnalysisMixin,
    EntryMixin, IOMixin, WindowMixin, CanvasMixin,
    ctk.CTk,
):
    def __init__(self): ...  # 상태 변수 초기화 + 빌드 호출
```

모든 Mixin은 `self`를 공유하는 Python 다중상속 패턴. 각 Mixin은 독립 클래스지만 실제로는 `MaintenanceApp` 인스턴스의 메서드로 동작함.

---

## 각 파일 담당 범위

| 파일 | 담당 |
|------|------|
| `constants.py` | `PX_PER_KM`, `METHOD_STYLES`, 색상 상수, `DIRECTIONS`, `LANES` 등 모든 전역 상수 |
| `utils.py` | `resource_path`, `get_user_data_dir`, `get_logger`, `load_method_settings`, `km_floor_to_grid`, `clamp`, `fmt_km`, `lane_conflict` 등 |
| `canvas_utils.py` | `_split_gaps`, `_merge_gaps`, `_draw_hatch_rect`, `_compute_overlap_status` |
| `mixins/ui_mixin.py` | `_build_ui`, `_build_dashboard`, `refresh_dashboard`, `render_legend`, `on_warranty_settings`, `on_manage_methods` 등 |
| `mixins/route_mixin.py` | `add_route`, `on_manage_routes`, `on_add_route`, `on_tab_changed`, `update_year_filter_values` 등 |
| `mixins/ic_mixin.py` | `on_manage_ic`, `on_manage_structures`, `_open_condition_data_dialog` |
| `mixins/analysis_mixin.py` | `on_manage_di/hpci/aar/rd/iri`, `on_improvement_priority`, `on_defect_risk`, `navigate_to_section` |
| `mixins/entry_mixin.py` | `on_add_entry`, `toggle_entry_mode`, `_apply_entry_mode_ui`, `generate_directions_from_name` |
| `mixins/io_mixin.py` | `on_save_csv`, `on_load_csv`, `on_load_excel`, `on_export_pdf`, `on_export_all_to_excel`, `_write_sqlite_cache`, `_load_from_sqlite_cache` |
| `mixins/window_mixin.py` | `on_closing`, `set_initial_window_size`, `auto_load_csvs_on_start` |
| `mixins/canvas_mixin.py` | `draw_schematic`, `draw_detail_schematic`, `toggle_magnifier_mode`, `open_entry_dialog`, `on_canvas_double_click`, `on_open_detail_table` 등 |
| `mixins/plan_mixin.py` | `on_business_plan`, `_bp_confirm`, `_inject_plan_entries`, `on_operation_plan_change`, `_build_hwp_payload`, `_run_hwp_export`, 사업계획/운영계획변경 저장·로드 |
| `hwpx_export.py` | `export_operation_change` — HWPX(zip+XML)를 직접 열어 제목/목적/표0·1·2 셀 채우기 (한글/COM 불필요) |

> 참고: `on_improvement_priority`(analysis_mixin)는 `apply_label`/`apply_callback`/`exclude_sections`/`default_year` 인자로 사업계획·운영계획변경 양쪽에서 재사용된다. 계획 엔트리는 `entries`에 `plan=True`로 주입되며, 이력 CSV·SQLite 저장 시 제외되고 `all_business_plan.csv`에 별도 보관 후 로드 시 재주입된다.

---

## 데이터 파일 (CSV)

프로그램 실행 폴더에 자동 저장/로드됨:

| 파일 | 내용 |
|------|------|
| `all_maintenance_history.csv` | 본선 포장 이력 |
| `all_branch_history.csv` | 램프/IC 포장 이력 |
| `all_ic.csv` | IC·JCT 위치 정보 |
| `all_structures.csv` | 교량·터널 정보 |
| `all_condition.csv` | DI·HPCI·AAR·RD·IRI 상태 데이터 |
| `all_aar_data.csv` | AAR 등급 데이터 |
| `all_routes_methods.csv` | 공법별 색상·카테고리 설정 |
| `warranty_settings.csv` | 카테고리별 하자기간 설정 |
| `all_business_plan.csv` | 사업계획 확정 구간 (모식도 '계획' 표기, 로드 시 entries에 재주입) |
| `all_operation_change.csv` | 운영계획변경 작성 구간 (비고=사업계획/운영계획 변경, orig 이정 포함) |
| `highway_data.db` | SQLite 캐시 (빠른 재로드용) |
