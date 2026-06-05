# -*- coding: utf-8 -*-
"""운영계획변경 한글(HWPX) 양식 자동 작성 모듈.

HWPX(=hwp+zip, OOXML 유사) 파일을 **순수 파이썬(zip + XML)** 으로 직접 채운다.
한글(HWP) 설치나 pywin32(COM)가 필요 없으며, Windows/리눅스 어디서나 동작한다.
결과 .hwpx 파일은 한글에서 그대로 열 수 있다.

템플릿: templates/operation_plan_template.hwpx
  표0 = 포장개량 세부위치 (9열 × 10행, 첫 2행 머리글, 3행=계, 4~10행=데이터 최대 7개소)
  표1 = 소요예산 산출내역 (5열 × 2행)
  표2 = 예산변경(안)     (5열 × 2행)
  제목 = 가로 원(rect) 안 2줄: 1줄 노선명(18pt), 2줄 사업명(22pt)
  목적 = 본문 단락

payload 키
  title_line1     : 제목 1줄(= 사용자 입력 노선명 전체)
  title_line2     : 제목 2줄(= 사업명)
  purpose         : 목적 문구
  business_content: 소요예산 산출내역 '사업내용'
  unit_price_str  : 단가 표기 문자열
  route_korean    : 표0 노선명 칸(한글 부분만, 예: '서산영덕선')
  rows            : [{loc, lane, length, init, change, delta, di, note}, ...] (최대 7)
  totals          : {count, length, init, change, delta}
"""
import copy
import io
import os
import zipfile
import xml.etree.ElementTree as ET

# ── 네임스페이스 ──────────────────────────────────────────────────────────
NS = {
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hp10": "http://www.hancom.co.kr/hwpml/2016/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hhs": "http://www.hancom.co.kr/hwpml/2011/history",
    "hm": "http://www.hancom.co.kr/hwpml/2011/master-page",
    "hpf": "http://www.hancom.co.kr/schema/2011/hpf",
    "dc": "http://purl.org/dc/elements/1.1/",
    "opf": "http://www.idpf.org/2007/opf/",
    "ooxmlchart": "http://www.hancom.co.kr/hwpml/2016/ooxmlchart",
    "hwpunitchar": "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar",
    "epub": "http://www.idpf.org/2007/ops",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
}
for _p, _u in NS.items():
    ET.register_namespace(_p, _u)

HP = "{%s}" % NS["hp"]
HH = "{%s}" % NS["hh"]

SECTION_XML = "Contents/section0.xml"
HEADER_XML = "Contents/header.xml"

# 제목 2줄째(사업명) 글꼴 크기(=22pt). 템플릿 charPr id와 목표 height(1/100 pt).
TITLE_LINE2_CHARPR = "25"
TITLE_LINE2_HEIGHT = "2200"

# 표0(포장개량 세부위치) 데이터 행 구성
TBL0_FIRST_DATA_ROW = 3   # 데이터 시작 rowAddr (0,1=머리글, 2=계)
TBL0_DATA_ROWS = 7        # 템플릿 기본 데이터 행 수(=7개소). 초과 시 자동 추가


# ── XML 헬퍼 ────────────────────────────────────────────────────────────
def _tc_addr(tc):
    addr = tc.find(HP + "cellAddr")
    if addr is None:
        return None, None
    return addr.get("colAddr"), addr.get("rowAddr")


def _find_cell(tbl, col, row):
    for tc in tbl.iter(HP + "tc"):
        c, r = _tc_addr(tc)
        if c == str(col) and r == str(row):
            return tc
    return None


def _set_run_text(run, text):
    """run의 첫 <hp:t>에 text를 넣고 나머지 텍스트 요소는 제거(서식=charPr 유지)."""
    ts = run.findall(HP + "t")
    if ts:
        first = ts[0]
        for sub in list(first):
            first.remove(sub)
        first.text = text
        first.tail = None
        for extra in ts[1:]:
            run.remove(extra)
    else:
        t = ET.SubElement(run, HP + "t")
        t.text = text


def _set_para_text(p, text):
    """단락의 첫 run만 남기고 text 설정(첫 run의 charPr=서식 유지)."""
    runs = p.findall(HP + "run")
    if not runs:
        run = ET.SubElement(p, HP + "run")
        _set_run_text(run, text)
        return
    _set_run_text(runs[0], text)
    for r in runs[1:]:
        p.remove(r)


def _set_cell_lines(tc, lines):
    """셀에 여러 줄(단락)을 설정. 줄 수에 맞춰 단락을 복제/삭제."""
    sub = tc.find(HP + "subList")
    if sub is None:
        return
    ps = sub.findall(HP + "p")
    if not ps:
        return
    lines = [("" if x is None else str(x)) for x in lines] or [""]
    # 부족하면 첫 단락 복제, 많으면 뒤 단락 삭제
    while len(ps) < len(lines):
        sub.append(copy.deepcopy(ps[0]))
        ps = sub.findall(HP + "p")
    for extra in ps[len(lines):]:
        sub.remove(extra)
    ps = sub.findall(HP + "p")
    for p, line in zip(ps, lines):
        _set_para_text(p, line)


def _set_cell(tc, text):
    _set_cell_lines(tc, [("" if text is None else str(text))])


def _iter_tables(root):
    return list(root.iter(HP + "tbl"))


# ── 본 작성 ──────────────────────────────────────────────────────────────
def _fill_title(root, payload):
    """가로 원(rect) 안 제목 2줄 채우기."""
    rect = None
    for r in root.iter(HP + "rect"):
        if r.find(".//" + HP + "drawText") is not None:
            rect = r
            break
    if rect is None:
        return
    sub = rect.find(".//" + HP + "drawText/" + HP + "subList")
    if sub is None:
        return
    ps = sub.findall(HP + "p")
    if len(ps) >= 1:
        _set_para_text(ps[0], payload.get("title_line1", ""))
    if len(ps) >= 2:
        _set_para_text(ps[1], payload.get("title_line2", ""))


def _fill_purpose(root, payload):
    purpose = payload.get("purpose", "")
    if not purpose:
        return
    # 표/도형 밖의 본문 단락 중 목적 문구를 가진 단락을 찾아 교체
    for tbl in root.iter(HP + "tbl"):
        pass
    target = None
    for p in root.iter(HP + "p"):
        txt = "".join(p.itertext())
        if "포장노후화" in txt or txt.strip().startswith("o "):
            target = p
            break
    if target is None:
        return
    lines = purpose.split("\n")
    runs = target.findall(HP + "run")
    if not runs:
        return
    # 첫 run의 서식을 유지하며 여러 줄 처리(줄바꿈은 lineBreak)
    base = runs[0]
    _set_run_text(base, lines[0])
    for r in runs[1:]:
        target.remove(r)
    for extra in lines[1:]:
        br = ET.SubElement(base, HP + "lineBreak")  # noqa: F841
        t = ET.SubElement(base, HP + "t")
        t.text = extra


def _fill_table0(tbl, payload):
    rows = payload.get("rows", [])
    totals = payload.get("totals", {})

    def put(col, row, val):
        tc = _find_cell(tbl, col, row)
        if tc is not None:
            _set_cell(tc, val)

    n = len(rows)
    total_data_rows = max(TBL0_DATA_ROWS, n)
    extra = total_data_rows - TBL0_DATA_ROWS

    # 개소가 7개를 넘으면 표준 데이터 행을 복제해 자동으로 행 추가
    if extra > 0:
        src = None  # 복제 원본: 일반 데이터 행(rowAddr=4, col1~col8)
        for tr in tbl.findall(HP + "tr"):
            tc = tr.find(HP + "tc")
            ca = tc.find(HP + "cellAddr") if tc is not None else None
            if ca is not None and ca.get("rowAddr") == str(TBL0_FIRST_DATA_ROW + 1):
                src = tr
                break
        children = list(tbl)
        tr_idxs = [i for i, c in enumerate(children) if c.tag == HP + "tr"]
        if src is not None and tr_idxs:
            last_tr_idx = max(tr_idxs)
            for k in range(extra):
                new_addr = TBL0_FIRST_DATA_ROW + TBL0_DATA_ROWS + k  # 10, 11, ...
                newtr = copy.deepcopy(src)
                for tc in newtr.findall(HP + "tc"):
                    ca = tc.find(HP + "cellAddr")
                    if ca is not None:
                        ca.set("rowAddr", str(new_addr))
                    _set_cell(tc, "")  # 복제 셀 비우기
                tbl.insert(last_tr_idx + 1 + k, newtr)
            try:
                tbl.set("rowCnt", str(int(tbl.get("rowCnt", "10")) + extra))
            except Exception:
                pass

    # 노선명 병합셀(row3 col0)의 rowSpan = 전체 데이터 행 수
    cell0 = _find_cell(tbl, 0, TBL0_FIRST_DATA_ROW)
    if cell0 is not None:
        span = cell0.find(HP + "cellSpan")
        if span is not None:
            span.set("rowSpan", str(total_data_rows))

    # 계 행(row2): col0='계'(유지), col1=개소수, col3=연장계, col4~6=사업비계
    put(1, 2, totals.get("count", ""))
    put(3, 2, totals.get("length", ""))
    put(4, 2, totals.get("init", ""))
    put(5, 2, totals.get("change", ""))
    put(6, 2, totals.get("delta", ""))

    # 노선명(병합셀, row3 col0) — 한글 부분만
    put(0, TBL0_FIRST_DATA_ROW, payload.get("route_korean", ""))

    # 데이터 행: col1~col8
    for i in range(total_data_rows):
        row = TBL0_FIRST_DATA_ROW + i
        r = rows[i] if i < len(rows) else None
        if r is None:
            vals = ["", "", "", "", "", "", "", ""]
        else:
            vals = [
                r.get("loc", ""), r.get("lane", ""), r.get("length", ""),
                r.get("init", ""), r.get("change", ""), r.get("delta", ""),
                r.get("di", ""), r.get("note", ""),
            ]
        for ci, v in enumerate(vals):
            put(1 + ci, row, v)


def _fill_table1(tbl, payload):
    totals = payload.get("totals", {})
    # row1: 사업내용, 사업량(=총연장), 단가, 공사비(=변경계), 비고(빈칸)
    for col, val in (
        (0, payload.get("business_content", "")),
        (1, totals.get("length", "")),
        (2, payload.get("unit_price_str", "")),
        (3, totals.get("change", "")),
    ):
        tc = _find_cell(tbl, col, 1)
        if tc is not None:
            _set_cell(tc, val)


def _fill_table2(tbl, payload):
    totals = payload.get("totals", {})
    # row1: 단위사업명(2줄), 당초계, 변경계, 증감계, 비고(빈칸)
    tc = _find_cell(tbl, 0, 1)
    if tc is not None:
        _set_cell_lines(tc, [payload.get("title_line1", ""),
                             payload.get("title_line2", "")])
    for col, val in (
        (1, totals.get("init", "")),
        (2, totals.get("change", "")),
        (3, totals.get("delta", "")),
    ):
        tc = _find_cell(tbl, col, 1)
        if tc is not None:
            _set_cell(tc, val)


def _patch_header_title_size(header_root):
    """제목 2줄째 charPr(id=25) 글꼴 크기를 22pt로 조정."""
    for c in header_root.iter(HH + "charPr"):
        if c.get("id") == TITLE_LINE2_CHARPR:
            c.set("height", TITLE_LINE2_HEIGHT)


def _xml_bytes(root):
    buf = io.BytesIO()
    ET.ElementTree(root).write(buf, encoding="UTF-8", xml_declaration=True)
    data = buf.getvalue()
    # 한글이 기대하는 선언 형태로 정규화
    if data.startswith(b"<?xml"):
        nl = data.find(b"?>")
        if nl != -1:
            data = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
                    + data[nl + 2:])
    return data


def export_operation_change(template_path, out_path, payload):
    """HWPX 템플릿을 읽어 payload로 채운 뒤 out_path에 저장."""
    with zipfile.ZipFile(template_path, "r") as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}
        infos = {n: zin.getinfo(n) for n in names}

    # section0.xml 채우기
    sec_root = ET.fromstring(contents[SECTION_XML])
    _fill_title(sec_root, payload)
    _fill_purpose(sec_root, payload)
    tbls = _iter_tables(sec_root)
    if len(tbls) >= 1:
        _fill_table0(tbls[0], payload)
    if len(tbls) >= 2:
        _fill_table1(tbls[1], payload)
    if len(tbls) >= 3:
        _fill_table2(tbls[2], payload)
    contents[SECTION_XML] = _xml_bytes(sec_root)

    # header.xml: 제목 2줄째 글꼴 크기 22pt
    try:
        hdr_root = ET.fromstring(contents[HEADER_XML])
        _patch_header_title_size(hdr_root)
        contents[HEADER_XML] = _xml_bytes(hdr_root)
    except Exception:
        pass

    # 다시 zip으로 저장 (mimetype은 첫 항목·무압축이어야 함)
    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    ordered = list(names)
    if "mimetype" in ordered:
        ordered.remove("mimetype")
        ordered.insert(0, "mimetype")

    with zipfile.ZipFile(out_path, "w") as zout:
        for n in ordered:
            data = contents[n]
            if n == "mimetype":
                zi = zipfile.ZipInfo("mimetype")
                zi.compress_type = zipfile.ZIP_STORED
                zout.writestr(zi, data)
            else:
                ct = getattr(infos.get(n), "compress_type", zipfile.ZIP_DEFLATED)
                if ct not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
                    ct = zipfile.ZIP_DEFLATED
                zout.writestr(n, data, compress_type=ct)
    return out_path
