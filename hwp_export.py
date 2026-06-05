# -*- coding: utf-8 -*-
"""운영계획변경 한글(HWP) 양식 자동 작성 모듈 (방법 A: 한글 COM 자동화).

요구 환경
─────────
- Windows + 한글(HWP)이 설치되어 있어야 함
- pywin32 패키지 필요:  pip install pywin32

동작 개요
─────────
1) 한글 COM 객체(HWPFrame.HwpObject) 생성, 보안 모듈 등록(보안 대화상자 억제)
2) 템플릿(operation_plan_template.hwp) 열기
3) 제목/목적 텍스트 치환, 3개 표(포장개량세부위치 / 소요예산 산출내역 / 예산변경안) 채우기
4) 다른 이름으로 저장

주의 (중요)
────────────
이 코드는 한글이 설치되지 않은 환경에서는 실행/검증이 불가능합니다.
표 셀 채우기는 '셀 단위 이동(TableRightCell)' 방식을 사용하며, 템플릿의 표 구조
(특히 포장개량세부위치 표의 '노선명' 세로 병합)에 따라 셀 인덱스 미세조정이
필요할 수 있습니다. 최초 1회는 한글 설치 PC에서 결과를 확인하세요.

포장개량세부위치 표는 노선명 칸이 세로 병합되어 있으면 자동 채우기 정렬이
어긋날 수 있습니다. 안정적인 자동 작성을 위해 템플릿에서 노선명 열의 세로 병합을
해제(각 행마다 별도 칸)해 두는 것을 권장합니다.
"""
import os


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


def _insert_text(hwp, text):
    act = hwp.HAction
    pset = hwp.HParameterSet.HInsertText
    act.GetDefault("InsertText", pset.HSet)
    pset.Text = str(text)
    act.Execute("InsertText", pset.HSet)


def _set_cell(hwp, text):
    """현재 셀의 내용을 text로 교체."""
    try:
        hwp.HAction.Run("TableSelCell")   # 현재 셀 블록 선택
        hwp.HAction.Run("Delete")         # 기존 내용 삭제
    except Exception:
        pass
    _insert_text(hwp, text)


def _next_cell(hwp):
    """오른쪽(행 끝이면 다음 행 첫 칸) 셀로 이동. 마지막 셀에서는 새 행 생성."""
    hwp.HAction.Run("TableRightCell")


def _enter_table(hwp, ctrl):
    """주어진 표 컨트롤의 첫 번째 셀로 진입."""
    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
    hwp.FindCtrl()
    hwp.HAction.Run("ShapeObjTableSelCell")  # 표 진입(첫 셀 선택)
    hwp.HAction.Run("Cancel")                # 선택 해제, 커서는 첫 셀 안


def _collect_tables(hwp):
    """문서 내 표 컨트롤을 문서 순서대로 수집."""
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


def _fill_title(hwp, payload):
    """제목/목적 텍스트를 찾아 치환(템플릿 예시 텍스트 기반)."""
    name = payload.get("project_name", "")
    purpose = payload.get("purpose", "")

    def _replace(find, repl):
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

    # 템플릿의 제목 예시를 사용자가 입력한 단위사업명으로 치환
    _replace("중앙선 246.9-313.5k", name)
    # 목적 문구가 다르면 치환(기본값과 동일하면 변화 없음)
    if purpose:
        _replace(
            "포장노후화 및 동절기 내 제설작업에 의한 노면파손이 발생함에 따라 "
            "안전한 주행환경 제공을 위해 포장개량 추진",
            purpose,
        )


def _fill_table1(hwp, ctrl, payload):
    """포장개량세부위치 표.

    구조(템플릿): 헤더 2행 + '계' 행 + 데이터 N행, 9열
      [노선명, 위치, 차로, 연장, 당초, 변경, 증감, DI등급, 비고]
    데이터 행은 가변. 셀 이동으로 '계' 행부터 채운다.
    """
    rows = payload.get("table1_rows", [])
    totals = payload.get("table1_totals", {})
    if not rows:
        return

    _enter_table(hwp, ctrl)

    # 헤더(상단 2행)를 시각 셀 기준으로 통과: 행1=7칸 + 행2(당초/변경/증감)=3칸 = 10칸.
    for _ in range(10):
        _next_cell(hwp)

    # '계' 행: 노선명='계', (위치 비움), 차로 비움, 연장=합계, 당초, 변경, 증감, DI 비움, 비고 비움
    sum_cells = ["계", totals.get("count", ""), "", totals.get("length", ""),
                 totals.get("init", ""), totals.get("change", ""), totals.get("delta", ""),
                 "", ""]
    for i, val in enumerate(sum_cells):
        _set_cell(hwp, val)
        if i < len(sum_cells) - 1:
            _next_cell(hwp)
    _next_cell(hwp)

    # 데이터 행들
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
        _next_cell(hwp)  # 다음 행 첫 칸(마지막 행이면 새 행 생성)


def _fill_simple_table(hwp, ctrl, header_count, data_rows):
    """단순 표(세로 병합 없음)용: 헤더 통과 후 데이터 행들을 순서대로 채움."""
    _enter_table(hwp, ctrl)
    for _ in range(header_count):
        _next_cell(hwp)
    for r in data_rows:
        for i, val in enumerate(r):
            _set_cell(hwp, val)
            if i < len(r) - 1:
                _next_cell(hwp)
        _next_cell(hwp)


def _fill_table2(hwp, ctrl, payload):
    """소요예산 산출내역 표 (헤더 5칸 + 데이터 행들)."""
    brows = payload.get("table2_rows", [])
    data = [[b.get("content", ""), b.get("qty", ""), b.get("unit_price", ""),
             b.get("cost", ""), b.get("note", "야간")] for b in brows]
    if data:
        _fill_simple_table(hwp, ctrl, 5, data)


def _fill_table3(hwp, ctrl, payload):
    """예산변경(안) 표 (헤더 5칸 + 데이터 1행)."""
    t3 = payload.get("table3", {})
    data = [[t3.get("name", ""), t3.get("init", ""), t3.get("change", ""),
             t3.get("delta", ""), ""]]
    _fill_simple_table(hwp, ctrl, 5, data)


def _save_as(hwp, out_path):
    act = hwp.HAction
    pset = hwp.HParameterSet.HFileOpenSave
    act.GetDefault("FileSaveAs_S", pset.HSet)
    pset.filename = out_path
    pset.Format = "HWP"
    act.Execute("FileSaveAs_S", pset.HSet)


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
        try:
            _fill_title(hwp, payload)
        except Exception:
            pass
        tables = _collect_tables(hwp)
        if len(tables) >= 1:
            try:
                _fill_table1(hwp, tables[0], payload)
            except Exception:
                pass
        if len(tables) >= 2:
            try:
                _fill_table2(hwp, tables[1], payload)
            except Exception:
                pass
        if len(tables) >= 3:
            try:
                _fill_table3(hwp, tables[2], payload)
            except Exception:
                pass
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
