# -*- coding: utf-8 -*-
"""운영계획변경 한글(HWP) 양식 자동 작성 모듈 (방법 A: 한글 COM 자동화).

요구 환경
─────────
- Windows + 한글(HWP)이 설치되어 있어야 함
- pywin32 패키지 필요:  pip install pywin32

채우기 방식 (중요)
──────────────────
이 모듈은 **누름틀(필드) 방식**으로 표를 채웁니다. 템플릿의 각 칸에 이름을 가진
누름틀(필드)을 심어두면, 셀 병합·머리글 구조와 무관하게 `PutFieldText(이름, 값)`
으로 정확히 그 칸에 값이 들어갑니다. (기존의 'TableRightCell 셀 이동' 방식은
세로/가로 병합 셀에서 정렬이 어긋나 값이 머리글로 침범하거나 여러 행이 한 칸에
뭉치는 문제가 있었습니다.)

템플릿에 심어야 할 필드 이름 (입력 → 누름틀, 속성에서 '필드 이름' 지정)
─────────────────────────────────────────────────────────────────────
[제목/목적]
  title                      ← 단위사업명(제목)
  purpose                    ← 목적 문구

[표1 · 포장개량 세부위치]  (※ 데이터 행 수는 템플릿에 만든 필드 개수만큼 자동 인식)
  계 행:  t1_total_count, t1_total_len, t1_total_init, t1_total_change, t1_total_delta
  i번째 데이터 행(1,2,3,…):
     t1_route_i, t1_loc_i, t1_lane_i, t1_len_i,
     t1_init_i, t1_change_i, t1_delta_i, t1_di_i, t1_note_i

[표2 · 소요예산 산출내역]  (i = 1,2,3,…)
     t2_content_i, t2_qty_i, t2_price_i, t2_cost_i, t2_note_i

[표3 · 예산변경(안)]
     t3_name, t3_init, t3_change, t3_delta, t3_note

규칙
─────
- 템플릿에 **존재하는 필드만** 채웁니다. 데이터가 필드보다 적으면 남는 칸은 빈칸,
  많으면 필드 개수까지만 채웁니다(초과분은 무시되며 경고 로그를 남깁니다).
- title/purpose 필드가 없으면 기존 찾아바꾸기로 제목/목적을 치환합니다.
- 문서에 우리 규칙의 필드가 하나도 없으면, 과거 호환을 위해 옛 셀-이동 방식으로
  대체 동작합니다(권장하지 않음 — 병합 양식에서 깨질 수 있음).
"""
import os


# ─────────────────────────────────────────────────────────────────────────
# 환경/객체
# ─────────────────────────────────────────────────────────────────────────
def hwp_available() -> bool:
    """한글 COM 자동화 사용 가능 여부."""
    if os.name != "nt":
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def _get_hwp():
    import win32com.client as win32
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    # 파일 접근 보안 대화상자 억제
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass
    return hwp


# ─────────────────────────────────────────────────────────────────────────
# 필드(누름틀) 헬퍼
# ─────────────────────────────────────────────────────────────────────────
def _field_exists(hwp, name) -> bool:
    try:
        return bool(hwp.FieldExist(name))
    except Exception:
        return False


def _put_field(hwp, name, text):
    """이름이 name인 누름틀에 text를 채운다. 없으면 조용히 건너뜀."""
    if not _field_exists(hwp, name):
        return False
    try:
        hwp.PutFieldText(name, "" if text is None else str(text))
        return True
    except Exception:
        return False


def _any_field_exists(hwp, names) -> bool:
    return any(_field_exists(hwp, n) for n in names)


# ─────────────────────────────────────────────────────────────────────────
# 제목/목적
# ─────────────────────────────────────────────────────────────────────────
def _replace_text(hwp, find, repl):
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HFindReplace
        act.GetDefault("AllReplace", pset.HSet)
        pset.FindString = find
        pset.ReplaceString = repl
        pset.IgnoreMessage = 1
        pset.Direction = 0
        pset.ReplaceMode = 1
        act.Execute("AllReplace", pset.HSet)
    except Exception:
        pass


def _fill_title(hwp, payload):
    name = payload.get("project_name", "")
    purpose = payload.get("purpose", "")

    used_field = False
    if name and _put_field(hwp, "title", name):
        used_field = True
    if purpose and _put_field(hwp, "purpose", purpose):
        used_field = True
    if used_field:
        return

    # 필드가 없으면 예시 텍스트를 찾아 치환(구버전 템플릿 호환)
    if name:
        _replace_text(hwp, "중앙선 246.9-313.5k", name)
    if purpose:
        _replace_text(
            hwp,
            "포장노후화 및 동절기 내 제설작업에 의한 노면파손이 발생함에 따라 "
            "안전한 주행환경 제공을 위해 포장개량 추진",
            purpose,
        )


# ─────────────────────────────────────────────────────────────────────────
# 표 채우기 (필드 방식)
# ─────────────────────────────────────────────────────────────────────────
def _count_row_fields(hwp, name_fmt, max_scan=200):
    """name_fmt.format(i)에 해당하는 필드가 연속으로 몇 개 존재하는지 센다(1부터)."""
    n = 0
    for i in range(1, max_scan + 1):
        if _field_exists(hwp, name_fmt.format(i)):
            n = i
        else:
            break
    return n


def _fill_rows(hwp, col_fmts, data_rows):
    """행 단위 필드 채우기.

    col_fmts: 컬럼별 필드 이름 포맷 리스트. 예) ["t1_route_{}", "t1_loc_{}", ...]
              각 포맷은 {}.format(행번호)로 실제 필드 이름이 된다.
    data_rows: [(c0, c1, ...), ...] 또는 [[...], ...] 형태의 행 값 리스트.

    템플릿에 만들어 둔 필드 개수만큼만 채우고, 남는 행 필드는 빈칸으로 비운다.
    데이터가 필드보다 많으면 초과분은 버린다(경고 로그).
    반환: 사용된(존재하는) 행 필드가 하나라도 있으면 True.
    """
    if not col_fmts:
        return False
    # 첫 컬럼 포맷 기준으로 템플릿의 데이터 행 개수 파악
    n_fields = _count_row_fields(hwp, col_fmts[0])
    if n_fields == 0:
        return False

    if len(data_rows) > n_fields:
        try:
            from utils import log_warning
            log_warning(
                f"HWP 양식 행 부족: 데이터 {len(data_rows)}행 > 템플릿 필드 {n_fields}행. "
                f"초과분은 누락됩니다. 템플릿 행/필드를 늘려주세요. (필드 예: {col_fmts[0]})"
            )
        except Exception:
            pass

    for ridx in range(1, n_fields + 1):
        values = data_rows[ridx - 1] if ridx - 1 < len(data_rows) else [""] * len(col_fmts)
        for cidx, fmt in enumerate(col_fmts):
            val = values[cidx] if cidx < len(values) else ""
            _put_field(hwp, fmt.format(ridx), val)
    return True


def _fill_table1_fields(hwp, payload):
    rows = payload.get("table1_rows", [])
    totals = payload.get("table1_totals", {})

    # 계(합계) 행
    _put_field(hwp, "t1_total_count", totals.get("count", ""))
    _put_field(hwp, "t1_total_len", totals.get("length", ""))
    _put_field(hwp, "t1_total_init", totals.get("init", ""))
    _put_field(hwp, "t1_total_change", totals.get("change", ""))
    _put_field(hwp, "t1_total_delta", totals.get("delta", ""))

    data = [
        [
            r.get("route", ""), r.get("location", ""), r.get("lane", ""),
            r.get("length", ""), r.get("init", "-"), r.get("change", ""),
            r.get("delta", ""), r.get("di", "5~7"), r.get("note", ""),
        ]
        for r in rows
    ]
    col_fmts = [
        "t1_route_{}", "t1_loc_{}", "t1_lane_{}", "t1_len_{}",
        "t1_init_{}", "t1_change_{}", "t1_delta_{}", "t1_di_{}", "t1_note_{}",
    ]
    return _fill_rows(hwp, col_fmts, data)


def _fill_table2_fields(hwp, payload):
    brows = payload.get("table2_rows", [])
    data = [
        [b.get("content", ""), b.get("qty", ""), b.get("unit_price", ""),
         b.get("cost", ""), b.get("note", "야간")]
        for b in brows
    ]
    col_fmts = ["t2_content_{}", "t2_qty_{}", "t2_price_{}", "t2_cost_{}", "t2_note_{}"]
    return _fill_rows(hwp, col_fmts, data)


def _fill_table3_fields(hwp, payload):
    t3 = payload.get("table3", {})
    used = False
    used |= _put_field(hwp, "t3_name", t3.get("name", ""))
    used |= _put_field(hwp, "t3_init", t3.get("init", ""))
    used |= _put_field(hwp, "t3_change", t3.get("change", ""))
    used |= _put_field(hwp, "t3_delta", t3.get("delta", ""))
    used |= _put_field(hwp, "t3_note", "")
    return used


# ─────────────────────────────────────────────────────────────────────────
# (구버전 호환) 셀-이동 방식 — 필드가 전혀 없을 때만 사용
# ─────────────────────────────────────────────────────────────────────────
def _insert_text(hwp, text):
    act = hwp.HAction
    pset = hwp.HParameterSet.HInsertText
    act.GetDefault("InsertText", pset.HSet)
    pset.Text = str(text)
    act.Execute("InsertText", pset.HSet)


def _set_cell(hwp, text):
    try:
        hwp.HAction.Run("TableSelCell")
        hwp.HAction.Run("Delete")
    except Exception:
        pass
    _insert_text(hwp, text)


def _next_cell(hwp):
    hwp.HAction.Run("TableRightCell")


def _enter_table(hwp, ctrl):
    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
    hwp.FindCtrl()
    hwp.HAction.Run("ShapeObjTableSelCell")
    hwp.HAction.Run("Cancel")


def _collect_tables(hwp):
    tables = []
    ctrl = hwp.HeadCtrl
    while ctrl:
        try:
            if ctrl.CtrlID == "tbl":
                tables.append(ctrl)
        except Exception:
            pass
        ctrl = ctrl.Next
    return tables


def _legacy_fill_table1(hwp, ctrl, payload):
    rows = payload.get("table1_rows", [])
    totals = payload.get("table1_totals", {})
    if not rows:
        return
    _enter_table(hwp, ctrl)
    for _ in range(10):
        _next_cell(hwp)
    sum_cells = ["계", totals.get("count", ""), "", totals.get("length", ""),
                 totals.get("init", ""), totals.get("change", ""), totals.get("delta", ""),
                 "", ""]
    for i, val in enumerate(sum_cells):
        _set_cell(hwp, val)
        if i < len(sum_cells) - 1:
            _next_cell(hwp)
    _next_cell(hwp)
    for r in rows:
        cells = [
            r.get("route", ""), r.get("location", ""), r.get("lane", ""),
            r.get("length", ""), r.get("init", "-"), r.get("change", ""),
            r.get("delta", ""), r.get("di", "5~7"), r.get("note", ""),
        ]
        for i, val in enumerate(cells):
            _set_cell(hwp, val)
            if i < len(cells) - 1:
                _next_cell(hwp)
        _next_cell(hwp)


def _legacy_fill_simple_table(hwp, ctrl, header_count, data_rows):
    _enter_table(hwp, ctrl)
    for _ in range(header_count):
        _next_cell(hwp)
    for r in data_rows:
        for i, val in enumerate(r):
            _set_cell(hwp, val)
            if i < len(r) - 1:
                _next_cell(hwp)
        _next_cell(hwp)


def _legacy_fill(hwp, payload):
    tables = _collect_tables(hwp)
    if len(tables) >= 1:
        try:
            _legacy_fill_table1(hwp, tables[0], payload)
        except Exception:
            pass
    if len(tables) >= 2:
        brows = payload.get("table2_rows", [])
        data = [[b.get("content", ""), b.get("qty", ""), b.get("unit_price", ""),
                 b.get("cost", ""), b.get("note", "야간")] for b in brows]
        if data:
            try:
                _legacy_fill_simple_table(hwp, tables[1], 5, data)
            except Exception:
                pass
    if len(tables) >= 3:
        t3 = payload.get("table3", {})
        data = [[t3.get("name", ""), t3.get("init", ""), t3.get("change", ""),
                 t3.get("delta", ""), ""]]
        try:
            _legacy_fill_simple_table(hwp, tables[2], 5, data)
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────
# 저장 / 진입점
# ─────────────────────────────────────────────────────────────────────────
def _save_as(hwp, out_path):
    act = hwp.HAction
    pset = hwp.HParameterSet.HFileOpenSave
    act.GetDefault("FileSaveAs_S", pset.HSet)
    pset.filename = out_path
    pset.Format = "HWP"
    act.Execute("FileSaveAs_S", pset.HSet)


# 문서에 우리 규칙의 필드가 하나라도 있는지 판별할 때 검사하는 대표 이름들
_PRIMARY_FIELDS = [
    "title", "t1_route_1", "t1_total_count",
    "t2_content_1", "t3_name",
]


def export_operation_change(template_path, out_path, payload):
    """운영계획변경 한글 양식 작성.

    payload = {
        project_name, purpose, unit_price,
        table1_rows[], table1_totals{}, table2_rows[], table3{}
    }
    """
    hwp = _get_hwp()
    try:
        hwp.Open(template_path, "HWP", "forceopen:true")

        use_fields = _any_field_exists(hwp, _PRIMARY_FIELDS)
        if use_fields:
            try:
                _fill_title(hwp, payload)
            except Exception:
                pass
            for fn in (_fill_table1_fields, _fill_table2_fields, _fill_table3_fields):
                try:
                    fn(hwp, payload)
                except Exception:
                    pass
        else:
            # 필드가 없는 구버전 템플릿: 옛 셀-이동 방식으로 대체
            try:
                from utils import log_warning
                log_warning(
                    "HWP 템플릿에 누름틀(필드)이 없어 구버전 셀-이동 방식으로 작성합니다. "
                    "정확한 정렬을 위해 템플릿에 필드를 추가하세요(hwp_export.py 상단 안내 참고)."
                )
            except Exception:
                pass
            try:
                _fill_title(hwp, payload)
            except Exception:
                pass
            _legacy_fill(hwp, payload)

        _save_as(hwp, out_path)
    finally:
        try:
            hwp.Clear(1)
        except Exception:
            pass
        try:
            hwp.Quit()
        except Exception:
            pass
