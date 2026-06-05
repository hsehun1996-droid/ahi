# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - 모식도 그리기, 돋보기, 이력 상세 다이얼로그
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import csv
import customtkinter as ctk
import math
import sqlite3
import json
import logging
from datetime import datetime
import tkinter.font as tkfont
import io
import ctypes
import tempfile

try:
    from PIL import Image, ImageOps, ImageTk
    from reportlab.pdfgen import canvas as reportlab_canvas
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.utils import ImageReader
    _PDF_LIBS_AVAILABLE = True
except ImportError:
    _PDF_LIBS_AVAILABLE = False
    ImageTk = None

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    _EXCEL_LIBS_AVAILABLE = True
except ImportError:
    _EXCEL_LIBS_AVAILABLE = False

try:
    from CTkMessagebox import CTkMessagebox
except ImportError:
    CTkMessagebox = None

# 개별 파일을 직접 실행하거나 작업 폴더가 달라도 프로젝트 루트(상위 폴더)의
# constants/utils 등을 찾을 수 있도록 sys.path 에 추가
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from constants import *
from utils import (
    get_logger, log_exception, log_warning,
    read_csv_with_encoding, get_user_data_dir, resource_path,
    load_method_settings, reset_method_settings,
    get_method_warranty_period, build_source_signature,
    km_floor_to_grid, km_ceil_to_grid, clamp, fmt_km,
    intervals_overlap, lane_conflict,
)
from canvas_utils import (
    _split_gaps, _merge_gaps, _draw_hatch_rect, _compute_overlap_status,
)


class CanvasMixin:
    def km_to_x(self, km: float) -> int:
        # 유지: 기존 API를 호출하는 코드가 남아 있을 수 있어 보존
        r_start = float(self.route_start_km.get())
        return int((km - r_start) * PX_PER_KM)

    def _collect_ic_list(self, route, r_start, r_end):
        """노선 내 r_start~r_end 구간의 IC/JCT 목록을 반환."""
        ic_list = []
        try:
            for ic in getattr(self, 'ics', []):
                ic_route = str(ic.get('route') or '')
                if ic_route and ic_route != route['name']:
                    continue
                km_ic = float(ic.get('km', 0.0))
                if r_start - 1e-6 <= km_ic <= r_end + 1e-6:
                    ic_list.append((km_ic, ic))
        except Exception:
            pass
        ic_list.sort(key=lambda t: t[0])
        return ic_list

    def _draw_entries(self, canvas, route, bar1_top, bar2_top, bar_h, lane_count, x_of_km,
                      gaps_px, selected_year, item_map=None,
                      clip_range=None, always_label=False):
        """이력(entry)을 캔버스에 렌더링. item_map이 주어지면 item_id→idx 매핑 저장."""
        visible_entries = []
        for idx, it in enumerate(route["entries"]):
            wd = str(it.get("work_date") or "")
            year = wd[:4] if len(wd) >= 4 else None
            if clip_range is not None:
                if not intervals_overlap(it["start"], it["end"], clip_range[0], clip_range[1]):
                    continue
            elif selected_year and selected_year != "모두" and year != selected_year:
                continue
            visible_entries.append((idx, it, year))

        # ── 하자기간 기반 중복 판정 ──────────────────────────────────────────
        # '모두' 보기(또는 clip_range 뷰)에서만 적용.
        # 특정 연도 필터 시에는 기존 픽셀 겹침 방식 유지(해당 연도 이력을 그대로 표시).
        use_warranty_logic = (selected_year == "모두") or (clip_range is not None)
        if use_warranty_logic:
            overlap_status = _compute_overlap_status(visible_entries)
            # 오래된 이력부터 그려 최신 이력이 위에 표시되도록 연도 오름차순 정렬
            visible_entries.sort(key=lambda x: (x[2] is None, x[2] or ""))
        else:
            overlap_status = None  # 특정 연도 뷰: 기존 방식

        # geoms: (sx1, sx2, y1, y2, idx, yr_int, expiry_yr)
        # yr_int=-1 은 연도 불명, expiry_yr = yr_int + 하자기간
        geoms = []

        for idx, it, year in visible_entries:
            # 하자기간 로직: 숨김 항목은 렌더링 건너뜀
            if use_warranty_logic and overlap_status.get(idx) == 'hidden':
                continue

            if clip_range is not None:
                xs = max(it["start"], clip_range[0]); xe = min(it["end"], clip_range[1])
                if xe <= xs: continue
                x1, x2 = x_of_km(xs), x_of_km(xe)
            else:
                x1, x2 = x_of_km(it["start"]), x_of_km(it["end"])
            if x2 <= x1: x2 = x1 + 2

            style = METHOD_STYLES.get(it["method"], {"fill": "#718096"})
            top = bar1_top if it["direction"] == route["directions"][0] else bar2_top

            if it["lane"] == "전차로":
                y1, y2 = top, top + bar_h
            else:
                try: n = int(it["lane"].replace("차로", ""))
                except (ValueError, AttributeError): n = 1
                n = max(1, min(n, lane_count))
                seg_h = bar_h / lane_count
                if top == bar1_top:
                    y1 = top + bar_h - n * seg_h + 2; y2 = top + bar_h - (n - 1) * seg_h - 2
                else:
                    y1 = top + (n - 1) * seg_h + 2; y2 = top + n * seg_h - 2

            # 하자기간 계산 (빗금 판정용)
            try:
                yr_int = int(year) if year else -1
            except (ValueError, TypeError):
                yr_int = -1
            if yr_int >= 0:
                warranty = get_method_warranty_period(it.get("method", ""))
                expiry_yr = yr_int + warranty
            else:
                expiry_yr = -1

            for sx1, sx2 in _split_gaps(gaps_px, x1, x2):
                if sx2 <= sx1: continue
                item_id = canvas.create_rectangle(sx1, y1, sx2, y2, fill=style["fill"], outline="", tags=("entry",))
                if (always_label or selected_year == "모두") and sx2 - sx1 > 30 and year:
                    try:
                        bg = style.get("fill", "#718096")
                        rv, gv, bv = int(bg[1:3], 16), int(bg[3:5], 16), int(bg[5:7], 16)
                        tc = "#FFFFFF" if (rv*299 + gv*587 + bv*114) / 1000 < 140 else "#000000"
                    except Exception:
                        tc = "#FFFFFF"
                    # 사업계획으로 확정된 구간은 연도 대신 '계획'으로 표기
                    label_txt = "계획" if it.get("plan") else f"'{year[2:]}"
                    canvas.create_text((sx1+sx2)/2, (y1+y2)/2, text=label_txt,
                                       fill=tc, font=(self.font_family, CANVAS_FONT_S), tags=("entry_label",))
                if item_map is not None:
                    item_map[item_id] = idx
                geoms.append((sx1, sx2, y1, y2, idx, yr_int, expiry_yr))

        # ── 빗금 그리기: 픽셀 교차 영역에만 표시 ──────────────────────────
        # 하자기간 로직: 교차 쌍 중 newer.year ≤ older.expiry_year 인 경우에만 빗금
        # 기존 방식(특정 연도 필터): 픽셀 겹침이면 무조건 빗금
        geoms.sort(key=lambda g: g[0])
        n = len(geoms)
        for i in range(n):
            sx1a, sx2a, y1a, y2a, _ia, yr_a, exp_a = geoms[i]
            for j in range(i + 1, n):
                sx1b, sx2b, y1b, y2b, _ib, yr_b, exp_b = geoms[j]
                if sx1b >= sx2a:
                    break
                xr = min(sx2a, sx2b)
                yt = max(y1a, y1b); yb = min(y2a, y2b)
                if xr <= sx1b or yb <= yt:
                    continue
                if use_warranty_logic:
                    # 연도 정보가 없으면 빗금 생략
                    if yr_a < 0 or yr_b < 0:
                        continue
                    # 오래된 이력 기준으로 새 이력이 하자기간 내에 있을 때만 빗금
                    if yr_a < yr_b:
                        should_hatch = (yr_b <= exp_a)
                    elif yr_b < yr_a:
                        should_hatch = (yr_a <= exp_b)
                    else:
                        should_hatch = True  # 같은 연도 겹침은 중복으로 처리
                    if not should_hatch:
                        continue
                _draw_hatch_rect(canvas, sx1b, yt, xr, yb)

        canvas.tag_raise("entry_label")

    def _draw_bar_bg(self, canvas, gaps_px, total_w, bar_h, bar1_top, bar2_top, lane_count):
        """바탕 바(분절), 차로 구분선, 중분대를 그리고 lane_y_positions를 반환."""
        margin = 20; right = total_w - margin
        for y_top in (bar1_top, bar2_top):
            seg = margin
            for a, b in gaps_px:
                if a > seg:
                    canvas.create_rectangle(seg, y_top, a, y_top + bar_h, fill=BACKGROUND_FILL, outline="")
                seg = b
            if seg < right:
                canvas.create_rectangle(seg, y_top, right, y_top + bar_h, fill=BACKGROUND_FILL, outline="")

        lane_y_positions = []
        for top in (bar1_top, bar2_top):
            for j in range(1, lane_count):
                y = top + bar_h * j / lane_count
                seg = margin
                for a, b in gaps_px:
                    if a > seg: canvas.create_line(seg, y, a, y, fill=LANE_BORDER)
                    seg = b
                if seg < right: canvas.create_line(seg, y, right, y, fill=LANE_BORDER)
                lane_y_positions.append(y)

        # 중분대: gap 영역 안에만 그림 (바 침범 금지)
        median_h = bar_h / lane_count
        actual_gap = bar2_top - (bar1_top + bar_h)
        draw_median_h = min(median_h, actual_gap / 2) if actual_gap > 0 else 0
        if draw_median_h > 0:
            for y1, y2 in [(bar1_top + bar_h, bar1_top + bar_h + draw_median_h),
                           (bar2_top - draw_median_h, bar2_top)]:
                seg = margin
                for a, b in gaps_px:
                    if a > seg: canvas.create_rectangle(seg, y1, a, y2, fill="#000000", outline="")
                    seg = b
                if seg < right: canvas.create_rectangle(seg, y1, right, y2, fill="#000000", outline="")

        return lane_y_positions

    def _draw_overlay(self, canvas, grid_positions, km_positions, lane_y_positions,
                      gaps_px, total_w, bar1_top, bar2_top, bar_h, km_line_color=None):
        """눈금선·차로선·gap 마감선을 최상단에 다시 그린다. (100m 점선은 호출자가 마지막에 직접 그린다.)"""
        if km_line_color is None:
            km_line_color = GRID_1KM_COLOR
        margin = 20; right = total_w - margin
        for x in km_positions:
            canvas.create_line(x, bar1_top, x, bar1_top + bar_h, fill=km_line_color, width=2)
            canvas.create_line(x, bar2_top, x, bar2_top + bar_h, fill=km_line_color, width=2)
        for y in lane_y_positions:
            seg = margin
            for a, b in gaps_px:
                if a > seg: canvas.create_line(seg, y, a, y, fill=LANE_BORDER)
                seg = b
            if seg < right: canvas.create_line(seg, y, right, y, fill=LANE_BORDER)
        for a, b in gaps_px:
            for top in (bar1_top, bar2_top):
                canvas.create_line(a, top, a, top + bar_h, fill=LANE_BORDER, width=2)
                canvas.create_line(b, top, b, top + bar_h, fill=LANE_BORDER, width=2)

    def _draw_structures(self, canvas, route, r_start, r_end, bar1_top, bar2_top, bar_h,
                          x_of_km, gaps_px=(), init_font_size=9, min_font_size=5, min_width=8,
                          label_collector=None, canvas_w=0, canvas_h=0, segment_index=0):
        """교량/터널 구조물을 캔버스에 렌더링."""
        # 같은 bar에 동일 km 구간이 중복 그려지는 것을 방지 (예: 양방향 + 특정방향 항목 동시 존재)
        drawn_ranges = {bar1_top: [], bar2_top: []}
        for struct in route.get("structures", []):
            s_km, e_km = struct['start'], struct['end']
            if abs(s_km - e_km) < 1e-6:
                e_km = s_km + 0.1
            if not intervals_overlap(s_km, e_km, r_start, r_end):
                continue
            s_clipped = max(s_km, r_start); e_clipped = min(e_km, r_end)
            if e_clipped <= s_clipped: continue

            x1 = x_of_km(s_clipped); x2 = x_of_km(e_clipped)
            is_tunnel = struct.get('type', '교량') == "터널"
            stroke_color = "#999999" if is_tunnel else "#55CCFF"
            stipple_pattern = "gray12"

            target_tops = []
            s_dir = struct.get('direction', '양방향').strip()
            r_dir1 = route['directions'][0].strip()
            r_dir2 = route['directions'][1].strip() if len(route['directions']) > 1 else ""
            if s_dir == '양방향':
                target_tops = [bar1_top, bar2_top]
            elif s_dir == r_dir1:
                target_tops = [bar1_top]
            elif r_dir2 and s_dir == r_dir2:
                target_tops = [bar2_top]
            else:
                if s_dir in r_dir1 or r_dir1 in s_dir: target_tops.append(bar1_top)
                if r_dir2 and (s_dir in r_dir2 or r_dir2 in s_dir): target_tops.append(bar2_top)
                if not target_tops: target_tops = [bar1_top, bar2_top]

            raw_name = struct['name'] if isinstance(struct['name'], str) else "".join(struct['name'])
            # 세로쓰기: 글자 하나씩 줄바꿈
            vertical_text = "\n".join(list(raw_name))
            cx = (x1 + x2) / 2
            # km 구간 키 (소수점 3자리 반올림으로 float 오차 흡수)
            range_key = (round(s_clipped, 3), round(e_clipped, 3))

            for top in target_tops:
                # 이미 같은 bar + 같은 구간에 그린 경우 건너뜀 (양방향/방향 혼재로 인한 중복 방지)
                if range_key in drawn_ranges[top]:
                    continue
                drawn_ranges[top].append(range_key)

                rect_kw = dict(outline=stroke_color, width=2, fill=stroke_color)
                if stipple_pattern:
                    rect_kw["stipple"] = stipple_pattern
                if gaps_px:
                    for sx1, sx2 in _split_gaps(gaps_px, x1, x2):
                        canvas.create_rectangle(sx1, top, sx2, top + bar_h, **rect_kw)
                else:
                    canvas.create_rectangle(x1, top, x2, top + bar_h, **rect_kw)
                cy = top + bar_h / 2
                char_count = len(raw_name)
                font_size = init_font_size; draw_text = False
                while font_size > min_font_size:
                    font = tkfont.Font(family=self.font_family, size=font_size, weight="bold")
                    char_h = font.metrics("linespace")
                    char_w = font.measure("가")  # 한글 한 글자 기준 폭
                    if char_count * char_h <= bar_h - 4 and char_w <= (x2 - x1) - 4:
                        draw_text = True; break
                    font_size -= 1
                if (x2 - x1) < min_width: draw_text = False
                if draw_text:
                    if label_collector is not None:
                        label_collector.append({
                            "type": "structure",
                            "text": vertical_text,
                            "cx": cx, "cy": cy,
                            "font_size": font_size,
                            "color": "black",
                            "segment": segment_index,
                            "canvas_w": canvas_w, "canvas_h": canvas_h,
                        })
                    else:
                        canvas.create_text(cx, cy, text=vertical_text, fill="black",
                                           font=(self.font_family, font_size, "bold"), anchor="center", justify="center")

    def _draw_condition_bands(self, canvas, route, bar1_top, bar2_top, bar_h, lane_count,
                               x_of_km, view_hpci, view_di, view_aar,
                               view_rd=False, view_iri=False,
                               r_start=None, r_end=None, font_family="", pdf_mode=False,
                               selected_year="모두"):
        """HPCI/DI/AAR/RD/IRI 등급을 각 차로 구간에 색상 띠로 렌더링."""
        num_views = view_hpci + view_di + view_aar + view_rd + view_iri
        if num_views == 0:
            return
        hpci_data = route.get('hpci_data', {})
        di_data   = route.get('di_data', {})
        aar_data  = route.get('aar_data', {})
        rd_data   = route.get('rd_data', {})
        iri_data  = route.get('iri_data', {})
        all_keys  = (set(hpci_data.keys()) | set(di_data.keys()) | set(aar_data.keys())
                     | set(rd_data.keys()) | set(iri_data.keys()))
        if not all_keys:
            return

        hpci_colors = {5: "#FEE2E2", 6: "#FCA5A5", 7: "#EF4444"}
        di_colors   = {5: "#BFDBFE", 6: "#60A5FA", 7: "#2563EB"}
        aar_colors  = {2: "#C6F6D5", 3: "#68D391", 4: "#38A169"}
        rd_colors   = {5: "#FEF3C7", 6: "#FCD34D", 7: "#F59E0B"}
        iri_colors  = {5: "#EDE9FE", 6: "#A78BFA", 7: "#7C3AED"}

        def _get_for_year(data_map):
            """연도 필터에 맞는 값을 반환합니다. 필터가 없으면 최신 연도."""
            if not data_map:
                return None
            if not isinstance(data_map, dict):
                return data_map
            if selected_year and selected_year != "모두" and selected_year in data_map:
                return data_map[selected_year]
            return data_map[sorted(data_map.keys())[-1]]

        for (d, l, skm_str) in all_keys:
            try:
                start_km = float(skm_str)
                end_km = start_km + GRID_KM
            except ValueError:
                continue
            if r_start is not None:
                if not intervals_overlap(start_km, end_km, r_start, r_end):
                    continue
                s = max(start_km, r_start); e = min(end_km, r_end)
                if e <= s: continue
                x1, x2 = x_of_km(s), x_of_km(e)
            else:
                x1, x2 = x_of_km(start_km), x_of_km(end_km)

            top = bar1_top if d == route["directions"][0] else bar2_top
            try:
                n = int(l.replace("차로", ""))
            except (ValueError, AttributeError):
                continue
            seg_h = bar_h / lane_count
            y1 = (top + (n - 1) * seg_h if top == bar2_top else top + bar_h - n * seg_h) + 2
            y2 = y1 + seg_h - 4

            single_view = (pdf_mode and num_views == 1)
            if num_views > 1:
                sub_h = (y2 - y1) / num_views
                font_size = 6
                vi = 0
                y_hpci_1 = y_hpci_2 = y_di_1 = y_di_2 = y_aar_1 = y_aar_2 = y_rd_1 = y_rd_2 = y_iri_1 = y_iri_2 = None
                if view_hpci:
                    y_hpci_1, y_hpci_2 = y1 + vi * sub_h, y1 + (vi + 1) * sub_h; vi += 1
                if view_di:
                    y_di_1, y_di_2 = y1 + vi * sub_h, y1 + (vi + 1) * sub_h; vi += 1
                if view_aar:
                    y_aar_1, y_aar_2 = y1 + vi * sub_h, y1 + (vi + 1) * sub_h; vi += 1
                if view_rd:
                    y_rd_1, y_rd_2 = y1 + vi * sub_h, y1 + (vi + 1) * sub_h; vi += 1
                if view_iri:
                    y_iri_1, y_iri_2 = y1 + vi * sub_h, y1 + (vi + 1) * sub_h; vi += 1
            else:
                y_hpci_1 = y_hpci_2 = y_di_1 = y_di_2 = y_aar_1 = y_aar_2 = y_rd_1 = y_rd_2 = y_iri_1 = y_iri_2 = None
                if view_hpci: y_hpci_1, y_hpci_2 = y1, y2
                if view_di:   y_di_1, y_di_2 = y1, y2
                if view_aar:  y_aar_1, y_aar_2 = y1, y2
                if view_rd:   y_rd_1, y_rd_2 = y1, y2
                if view_iri:  y_iri_1, y_iri_2 = y1, y2
                font_size = 9

            fs = (font_family, font_size, "normal")

            if view_hpci and y_hpci_1 is not None:
                g = _get_for_year(hpci_data.get((d, l, skm_str)))
                if g is not None:
                    if g in hpci_colors:
                        canvas.create_rectangle(x1+2, y_hpci_1, x2-2, y_hpci_2, fill=hpci_colors[g], outline="")
                    if g in hpci_colors or single_view:
                        canvas.create_text((x1+x2)/2, (y_hpci_1+y_hpci_2)/2, text=str(g),
                                           fill="#000000", font=fs, tags=("hpci_label",))
            if view_di and y_di_1 is not None:
                v = _get_for_year(di_data.get((d, l, skm_str)))
                if v is not None:
                    g = int(v); cg = min(g, 7)
                    if cg in di_colors:
                        canvas.create_rectangle(x1+2, y_di_1, x2-2, y_di_2, fill=di_colors[cg], outline="")
                    if cg in di_colors or single_view:
                        canvas.create_text((x1+x2)/2, (y_di_1+y_di_2)/2, text=str(g),
                                           fill="#000000", font=fs, tags=("di_label",))
            if view_aar and y_aar_1 is not None:
                g = _get_for_year(aar_data.get((d, l, skm_str)))
                if g is not None:
                    if g in aar_colors:
                        canvas.create_rectangle(x1+2, y_aar_1, x2-2, y_aar_2, fill=aar_colors[g], outline="")
                    if g in aar_colors or single_view:
                        canvas.create_text((x1+x2)/2, (y_aar_1+y_aar_2)/2, text=str(g),
                                           fill="#000000", font=fs, tags=("aar_label",))
            if view_rd and y_rd_1 is not None:
                v = _get_for_year(rd_data.get((d, l, skm_str)))
                if v is not None:
                    try:
                        g = int(float(v)); cg = min(g, 7)
                    except (ValueError, TypeError):
                        cg = None
                    if cg is not None:
                        if cg in rd_colors:
                            canvas.create_rectangle(x1+2, y_rd_1, x2-2, y_rd_2, fill=rd_colors[cg], outline="")
                        if cg in rd_colors or single_view:
                            canvas.create_text((x1+x2)/2, (y_rd_1+y_rd_2)/2, text=str(g),
                                               fill="#000000", font=fs, tags=("rd_label",))
            if view_iri and y_iri_1 is not None:
                v = _get_for_year(iri_data.get((d, l, skm_str)))
                if v is not None:
                    try:
                        g = int(float(v)); cg = min(g, 7)
                    except (ValueError, TypeError):
                        cg = None
                    if cg is not None:
                        if cg in iri_colors:
                            canvas.create_rectangle(x1+2, y_iri_1, x2-2, y_iri_2, fill=iri_colors[cg], outline="")
                        if cg in iri_colors or single_view:
                            canvas.create_text((x1+x2)/2, (y_iri_1+y_iri_2)/2, text=str(g),
                                               fill="#000000", font=fs, tags=("iri_label",))

        canvas.tag_raise("hpci_label")
        canvas.tag_raise("di_label")
        canvas.tag_raise("aar_label")
        canvas.tag_raise("rd_label")
        canvas.tag_raise("iri_label")

    def draw_schematic(self):
        # 요약 통계 항상 갱신 (노선 없을 때도)
        self.refresh_dashboard()
        # 현재 탭/노선 컨텍스트
        if self.current_route_index < 0 or self.current_route_index >= len(self.routes):
            return
        route = self.routes[self.current_route_index]
        canvas = route["canvas"]
        canvas.delete("all")
        
        # 보기 모드가 '기본'이 아니면(즉, 포장상태 항목 중 하나라도 켜지면) view_mode를 'CONDITION'으로 설정
        if self.view_hpci.get() or self.view_di.get() or self.view_aar.get() or self.view_rd.get() or self.view_iri.get():
            self.view_mode.set("CONDITION")
        
        r_start = float(route["start_km"]); r_end = float(route["end_km"])
        span_km = max(0.1, r_end - r_start)

        # 연도 필터 준비
        selected_year = self.filter_year.get()

        # IC 갭 정의 수집 (노선/구간 내) 및 총 추가 폭 계산
        ic_gap_w = IC_BOX_W + IC_GAP_MARGIN * 2
        ic_list = self._collect_ic_list(route, r_start, r_end)
        total_gaps_px = ic_gap_w * len(ic_list)

        total_w = int(span_km * PX_PER_KM) + 80 + total_gaps_px  # 갭 폭만큼 우측 여백 확대
        total_h = int(canvas.winfo_height()) or 360

        # 레이아웃 (display scale에 따라 constants.py에서 자동 조정)
        top_margin = CANVAS_TOP_MARGIN
        bar_h      = CANVAS_BAR_H
        dir_gap    = CANVAS_DIR_GAP
        bar1_top = top_margin
        bar2_top = bar1_top + bar_h + dir_gap
        bar_bottom = bar2_top + bar_h

        # 캔버스 스크롤영역
        canvas.config(scrollregion=(0, 0, total_w, max(total_h, bar_bottom + 80)))

        # km→px 매핑(갭을 실제 길이로 반영: 갭 위치 이후를 오른쪽으로 밀어냄)
        km_marks = [km for km, _ in ic_list]
        def offset_for(km_value: float) -> int:
            cnt = 0
            for gk in km_marks:
                if km_value >= gk - 1e-9:
                    cnt += 1
            return cnt * ic_gap_w

        def x_of_km(km_value: float) -> int:
            return int((km_value - r_start) * PX_PER_KM) + 20 + offset_for(km_value)

        # 돋보기 역산을 위한 km 매핑 정보 저장
        route["_km_mapping"] = {
            "r_start": r_start, "r_end": r_end,
            "km_marks": list(km_marks), "ic_gap_w": ic_gap_w,
            "px_per_km": PX_PER_KM,
        }
        # 강조 하이라이트 계산용 바 레이아웃 저장
        route["_bar_layout"] = {
            "bar1_top": bar1_top, "bar2_top": bar2_top, "bar_h": bar_h,
        }

        # IC 기반 gap 영역의 좌/우 픽셀 범위 계산 (갭은 해당 km 지점 '앞'에 삽입)
        gaps_px = []  # [(left,right)]
        for km_ic, _ic in ic_list:
            right_px = x_of_km(km_ic)
            left_px = right_px - ic_gap_w
            gaps_px.append((max(20, left_px), min(total_w - 20, right_px)))
        gaps_px = _merge_gaps(sorted(gaps_px))
        # 모식도 선택 모드 오버레이용으로 IC 갭 픽셀 범위 저장
        route["_gaps_px"] = gaps_px

        def in_gap(xpx: int) -> bool:
            for a, b in gaps_px:
                if a <= xpx <= b:
                    return True
            return False

        # 100m 그리드 위치 수집 (gap 구간은 생략)
        x = km_ceil_to_grid(r_start, GRID_KM)
        grid_positions = []
        while x <= r_end + 1e-9:
            xx = x_of_km(x)
            if not in_gap(xx):
                grid_positions.append(xx)
            x += GRID_KM

        # 1km 눈금/라벨 (세로선은 바 내부만) + 위치 저장 (gap 내부는 생략)
        km_positions = []
        km_tick = math.ceil(r_start)  # 다음 정수 km
        while km_tick <= r_end + 1e-6:
            x = x_of_km(km_tick)
            if not in_gap(x):
                # 위/아래 바 내부에만 진한 세로선 (검은 실선, 약간 두껍게)
                canvas.create_line(x, bar1_top, x, bar1_top + bar_h, fill=GRID_1KM_COLOR, width=2)
                canvas.create_line(x, bar2_top, x, bar2_top + bar_h, fill=GRID_1KM_COLOR, width=2)
                canvas.create_text(x, bar1_top - CANVAS_KM_LBL_OFFSET, text=f"{km_tick:.0f}k", fill=TEXT_COLOR, font=(self.font_family, CANVAS_FONT_M))
                canvas.create_text(x, bar2_top + bar_h + CANVAS_KM_LBL_OFFSET, text=f"{km_tick:.0f}k", fill=TEXT_COLOR, font=(self.font_family, CANVAS_FONT_M))
                km_positions.append(x)
            km_tick += 1.0

        # 바탕 바 / 차로선 / 중분대 그리기
        lane_count = route.get("lane_count", 4)
        lane_y_positions = self._draw_bar_bg(canvas, gaps_px, total_w, bar_h, bar1_top, bar2_top, lane_count)

        # HPCI, DI, AAR, RD, IRI 등급 그리기 (연도 필터 반영)
        self._draw_condition_bands(
            canvas, route, bar1_top, bar2_top, bar_h, lane_count, x_of_km,
            self.view_hpci.get(), self.view_di.get(), self.view_aar.get(),
            view_rd=self.view_rd.get(), view_iri=self.view_iri.get(),
            font_family=self.font_family, selected_year=selected_year)
        # 방향 라벨 (스크롤해도 고정 위치 유지)
        dir_lbl1 = canvas.create_text(24, bar1_top - CANVAS_DIR_LBL_OFFSET, anchor="w", text=f"{route['directions'][0]}", fill=TEXT_COLOR, font=(self.font_family, CANVAS_FONT_L, "bold"), tags="dir_label")
        dir_lbl2 = canvas.create_text(24, bar2_top + bar_h + CANVAS_DIR_LBL_OFFSET, anchor="w", text=f"{route['directions'][1]}", fill=TEXT_COLOR, font=(self.font_family, CANVAS_FONT_L, "bold"), tags="dir_label")
        _dir_label_ys = (bar1_top - 36, bar2_top + bar_h + 36)
        _dir_label_ids = (dir_lbl1, dir_lbl2)

        def _pin_dir_labels(*_args):
            x = canvas.canvasx(0) + 4
            for lid, ly in zip(_dir_label_ids, _dir_label_ys):
                canvas.coords(lid, x, ly)
            canvas.tag_raise("dir_label")

        hbar = route.get("hbar")
        if hbar:
            canvas.configure(xscrollcommand=lambda *a: (hbar.set(*a), _pin_dir_labels()))
        canvas.bind("<Configure>", _pin_dir_labels)

        # IC 마커 및 패널(콜아웃) 그리기
        try:
            last_ramp_bottom = bar_bottom
            for ic_idx_global, ic in enumerate(getattr(self, 'ics', [])):
                # 노선별 필터
                ic_route = str(ic.get('route') or '')
                if ic_route and ic_route != route['name']:
                    continue
                km = float(ic.get('km', 0.0))
                # 각 IC를 그리기 전에 last_ramp_bottom을 현재 IC의 박스 하단으로 초기화
                last_ramp_bottom = bar2_top + bar_h
                if km < r_start - 1e-6 or km > r_end + 1e-6:
                    continue
                # x 좌표(갭 반영 매핑)
                right_edge = x_of_km(km)
                left_edge = right_edge - ic_gap_w
                box_left = left_edge + IC_GAP_MARGIN
                box_right = right_edge - IC_GAP_MARGIN

                # 박스 자체 그리기
                box_top = bar1_top
                box_bottom = bar2_top + bar_h
                canvas.create_rectangle(box_left, box_top, box_right, box_bottom, outline=IC_COLOR, width=3, fill="#FFFFFF")

                # 박스 내부 텍스트(세로 쓰기)
                center_x = int((box_left + box_right) / 2)
                
                # 이름 및 이정 (세로 쓰기, 중앙 정렬)
                name = str(ic.get('name') or '')
                km_str = f"{km:.1f}k"
                char_h = CANVAS_IC_CHAR_H
                gap = max(6, char_h // 2)
                total_text_h = (len(name) + len(km_str)) * char_h + gap
                box_h = box_bottom - box_top
                y_cursor = box_top + (box_h - total_text_h) / 2

                for char in name:
                    canvas.create_text(center_x, y_cursor + char_h/2, text=char, fill=IC_COLOR, font=(self.font_family, CANVAS_FONT_XL, "bold"), anchor="center")
                    y_cursor += char_h

                y_cursor += gap

                for char in km_str:
                    canvas.create_text(center_x, y_cursor + char_h/2, text=char, fill=IC_COLOR, font=(self.font_family, CANVAS_FONT_L, "bold"), anchor="center")
                    y_cursor += char_h

                # 램프 미니 모식도 그리기 제거 (요청에 따라 램프 이력 미사용)

                # 박스 좌우에 100m 간격 강조 눈금 표시
                tick_start = max(r_start, km - IC_TICK_SPAN_KM)
                tick_end = min(r_end, km + IC_TICK_SPAN_KM)
                t = km_floor_to_grid(tick_start, GRID_KM)
                # 요청에 따라 강조 눈금(굵은 빨간선) 제거
            # 램프 모식도까지 포함한 높이로 스크롤 영역 보정
            canvas.config(scrollregion=(0, 0, total_w, max(total_h, last_ramp_bottom + 80)))
        except Exception:
            log_exception(f"IC 렌더링 실패: route={route.get('name', '')}")

        # 이력 그리기 (+ 캔버스 아이템과 이력 인덱스 매핑, 하단 리스트 채우기)
        route["canvas_item_to_index"].clear()
        route.setdefault("di_target_item_to_info", {}).clear()
        view_mode = self.view_mode.get()
        # 하단 리스트 갱신 준비
        tree = route.get("list_tree") # 이 변수는 현재 None으로 설정되어 사용되지 않습니다.
        if tree is not None:
            for item in tree.get_children():
                tree.delete(item)
            route["tree_item_to_index"].clear()

        # 이력 그리기
        if view_mode == "기본":
            self._draw_entries(canvas, route, bar1_top, bar2_top, bar_h, lane_count, x_of_km,
                               gaps_px, selected_year, item_map=route["canvas_item_to_index"])

        # 구조물 그리기 (교량, 터널) - 이력 위에 표시되도록 이력 그리기 이후에 호출
        self._draw_structures(canvas, route, r_start, r_end, bar1_top, bar2_top, bar_h,
                              x_of_km, gaps_px=gaps_px)
        
        # 오버레이: 1km 눈금선·차로선·gap 마감선을 맨 위로 다시 그림
        self._draw_overlay(canvas, grid_positions, km_positions, lane_y_positions,
                           gaps_px, total_w, bar1_top, bar2_top, bar_h)
        # 100m 점선 그리드: 모든 그리기 작업이 끝난 후 마지막에 직접 그려 최상단 보장
        try:
            for xx in grid_positions:
                canvas.create_line(xx, bar1_top, xx, bar1_top + bar_h, fill=GRID_100M, dash=(2, 4))
                canvas.create_line(xx, bar2_top, xx, bar2_top + bar_h, fill=GRID_100M, dash=(2, 4))
        except Exception:
            log_exception(f"그리드 렌더링 실패: route={route.get('name', '')}")

        # 모식도 선택 모드: 선택불가/선택됨 오버레이를 최상단에 그림
        if getattr(self, "schematic_select_mode", False):
            try:
                self._draw_schematic_selection_overlay(route)
            except Exception:
                log_exception("모식도 선택 오버레이 렌더링 실패")

        route["_last_render_signature"] = self._route_render_signature(route)
    # ───────────────── 돋보기 모드 ─────────────────

    # ───────────────── 구간 하이라이트 ─────────────────

    def _km_to_canvas_x(self, route, km: float) -> int:
        """km 값을 메인 캔버스 픽셀 x 좌표로 변환 (IC 갭 오프셋 포함)."""
        m = route.get("_km_mapping", {})
        r_start  = m.get("r_start", 0.0)
        km_marks = m.get("km_marks", [])
        ic_gap_w = m.get("ic_gap_w", 0)
        px_per_km = m.get("px_per_km", PX_PER_KM)
        cnt = sum(1 for gk in km_marks if km >= gk - 1e-9)
        return int((km - r_start) * px_per_km) + 20 + cnt * ic_gap_w

    def _start_magnifier_highlight(self, route: dict, sec_start: float, sec_end: float):
        """메인 캔버스에 sec_start~sec_end 구간 노란 하이라이트(깜빡임) 표시."""
        self._stop_magnifier_highlight(route)  # 기존 것 제거

        canvas = route.get("canvas")
        bl = route.get("_bar_layout", {})
        if not canvas or not bl:
            return

        x1 = self._km_to_canvas_x(route, sec_start)
        x2 = self._km_to_canvas_x(route, sec_end)
        y1 = bl["bar1_top"] - 4
        y2 = bl["bar2_top"] + bl["bar_h"] + 4

        # 반투명 노란 채움 + 굵은 테두리
        fill_id = canvas.create_rectangle(x1, y1, x2, y2,
                                          fill="#FFFF00", outline="",
                                          stipple="gray25", tags="mag_highlight")
        border_id = canvas.create_rectangle(x1, y1, x2, y2,
                                            fill="", outline="#FFD700",
                                            width=3, tags="mag_highlight")

        route["_highlight_ids"] = (fill_id, border_id)
        route["_highlight_visible"] = True
        route["_highlight_after"] = None
        self._blink_highlight(route)

    def _blink_highlight(self, route: dict):
        """하이라이트 깜빡임 애니메이션 (400ms 간격)."""
        canvas = route.get("canvas")
        ids = route.get("_highlight_ids")
        if not canvas or not ids:
            return
        visible = route.get("_highlight_visible", True)
        new_state = "hidden" if visible else "normal"
        for item_id in ids:
            canvas.itemconfigure(item_id, state=new_state)
        route["_highlight_visible"] = not visible
        after_id = canvas.after(400, lambda: self._blink_highlight(route))
        route["_highlight_after"] = after_id

    def _stop_magnifier_highlight(self, route: dict):
        """하이라이트 제거 및 애니메이션 취소."""
        canvas = route.get("canvas")
        after_id = route.pop("_highlight_after", None)
        if after_id and canvas:
            try:
                canvas.after_cancel(after_id)
            except Exception:
                pass
        if canvas:
            canvas.delete("mag_highlight")
        route.pop("_highlight_ids", None)
        route.pop("_highlight_visible", None)

    def _create_magnifier_cursor_path(self):
        """32×32 돋보기 모양 .cur 파일 생성 후 경로 반환. 실패 시 None."""
        if self._magnifier_cur_path and os.path.exists(self._magnifier_cur_path):
            return self._magnifier_cur_path
        try:
            import struct
            W, H = 32, 32
            cx, cy = 11, 11   # 렌즈 중심 (핫스팟)
            # 렌즈 원 (두께 2): r_in=7 ~ r_out=9
            px = set()
            for row in range(H):
                for col in range(W):
                    d = math.sqrt((col - cx) ** 2 + (row - cy) ** 2)
                    if 7 <= d <= 9:
                        px.add((col, row))
            # 손잡이: (19,19)→(27,27) 대각선, 굵기 2
            for i in range(9):
                for dx in range(2):
                    for dy in range(2):
                        hx, hy = 19 + i + dx, 19 + i + dy
                        if 0 <= hx < W and 0 <= hy < H:
                            px.add((hx, hy))

            # XOR mask (검정=0) / AND mask (불투명=0, 투명=1)
            xor_rows, and_rows = [], []
            for row in range(H):
                xv, av = 0, 0xFFFFFFFF
                for col in range(W):
                    if (col, row) in px:
                        av &= ~(1 << (31 - col))  # AND bit → 0 (불투명 검정)
                xor_rows.append(xv)
                and_rows.append(av)

            # BMP는 하단 행 먼저 저장
            def pack_rows(rows):
                return b''.join(struct.pack('>I', rows[r]) for r in range(H - 1, -1, -1))

            xor_data = pack_rows(xor_rows)
            and_data = pack_rows(and_rows)

            # BITMAPINFOHEADER (biHeight = H*2: XOR+AND 통합)
            bih = struct.pack('<IiiHHIIiiII',
                              40, W, H * 2, 1, 1, 0, 0, 0, 0, 0, 0)
            color_table = b'\x00\x00\x00\x00\xFF\xFF\xFF\x00'
            image_data = bih + color_table + xor_data + and_data

            # ICONDIR + ICONDIRENTRY
            icondir = struct.pack('<HHH', 0, 2, 1)
            icondirentry = struct.pack('<BBBBHHII',
                                      W, H, 0, 0,
                                      cx, cy,
                                      len(image_data), 6 + 16)
            cur_bytes = icondir + icondirentry + image_data

            tmp = tempfile.NamedTemporaryFile(suffix='.cur', delete=False)
            tmp.write(cur_bytes)
            tmp.close()
            self._magnifier_cur_path = tmp.name
            return self._magnifier_cur_path
        except Exception:
            log_exception("돋보기 커서 생성 실패")
            return None

    def toggle_magnifier_mode(self, event=None):
        """돋보기 모드 토글 (단축키 o). 텍스트 입력 위젯에 포커스 중이면 무시."""
        # 모식도 선택 모드 중에는 돋보기 토글 무시 (마우스 바인딩 충돌 방지)
        if getattr(self, "schematic_select_mode", False):
            return
        try:
            focused = self.focus_get()
            if focused and focused.winfo_class() in ('Entry', 'Text', 'TEntry'):
                return
        except Exception:
            pass

        self.magnifier_mode = not self.magnifier_mode
        if self.magnifier_mode:
            cur_path = self._create_magnifier_cursor_path()
            if cur_path:
                cursor = "@" + cur_path.replace("\\", "/")
            else:
                cursor = "crosshair"
        else:
            cursor = ""

        for route in self.routes:
            c = route.get("canvas")
            if not c:
                continue
            c.configure(cursor=cursor)
            if self.magnifier_mode:
                c.bind("<Button-1>", self.on_magnifier_click)
                c.bind("<Motion>", self._on_magnifier_motion)
                c.bind("<Leave>", self._hide_magnifier_tooltip)
            else:
                # 포커스 바인딩만 복원
                c.bind("<Button-1>", lambda e: e.widget.focus_set())
                c.unbind("<Motion>")
                c.unbind("<Leave>")

        if not self.magnifier_mode:
            self._hide_magnifier_tooltip()
            for route in self.routes:
                self._stop_magnifier_highlight(route)
            if self._magnifier_cur_path and os.path.exists(self._magnifier_cur_path):
                try:
                    os.remove(self._magnifier_cur_path)
                except Exception:
                    pass
                self._magnifier_cur_path = None

    def on_detail_canvas_double_click(self, event):
        """돋보기 상세 모식도에서 더블클릭 → 지사 이력 상세 다이얼로그."""
        widget = event.widget
        # 클릭된 캔버스가 어느 route의 detail_canvas인지 찾기
        route = None
        for r in self.routes:
            if r.get("detail_canvas") is widget:
                route = r
                break
        if route is None:
            return
        cx = widget.canvasx(event.x)
        cy = widget.canvasy(event.y)
        items = widget.find_overlapping(cx, cy, cx, cy)
        mapping = route.get("branch_item_to_index", {})
        hit_indices = [mapping[i] for i in items if i in mapping]
        if not hit_indices:
            return
        idx = hit_indices[-1]
        self.open_branch_entry_dialog(route, idx)

    def open_branch_entry_dialog(self, route: dict, idx: int):
        """지사 이력 상세/수정 다이얼로그."""
        entries = route.get("branch_entries", [])
        if idx < 0 or idx >= len(entries):
            return
        it = entries[idx]

        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.attributes("-topmost", True)

        var_start  = tk.StringVar(value=str(it.get("start", "")))
        var_end    = tk.StringVar(value=str(it.get("end", "")))
        var_method = tk.StringVar(value=it.get("method", "단면보수"))
        var_dir    = tk.StringVar(value=it.get("direction", ""))
        var_lane   = tk.StringVar(value=it.get("lane", "전차로"))
        var_date   = tk.StringVar(value=str(it.get("work_date", "")))

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=16)

        rows = [
            ("시작(km)", var_start), ("끝(km)", var_end),
            ("공법", var_method), ("시공날짜", var_date),
            ("방향", var_dir), ("차로", var_lane),
        ]
        for i, (lbl, var) in enumerate(rows):
            ctk.CTkLabel(body, text=lbl, anchor="w").grid(
                row=i, column=0, sticky="w", padx=(0, 8), pady=3)
            if lbl == "공법":
                self._create_styled_combobox(body, variable=var,
                                             values=list(BRANCH_METHOD_STYLES.keys()),
                                             width=160, height=28, state="readonly").grid(
                    row=i, column=1, sticky="ew", pady=3)
            elif lbl == "방향":
                self._create_styled_combobox(body, variable=var,
                                             values=route.get("directions", DIRECTIONS),
                                             width=160, height=28, state="readonly").grid(
                    row=i, column=1, sticky="ew", pady=3)
            elif lbl == "차로":
                self._create_styled_combobox(body, variable=var, values=LANES,
                                             width=160, height=28, state="readonly").grid(
                    row=i, column=1, sticky="ew", pady=3)
            else:
                ctk.CTkEntry(body, textvariable=var, width=160, height=28,
                             border_color="#404040").grid(
                    row=i, column=1, sticky="ew", pady=3)

        def do_save():
            try:
                s = float(var_start.get().strip())
                e = float(var_end.get().strip())
            except ValueError:
                self._show_error("입력 오류", "이정을 숫자로 입력해 주세요.")
                return
            if e <= s:
                self._show_error("입력 오류", "끝 이정이 시작 이정보다 커야 합니다.")
                return
            wd = var_date.get().strip()
            if len(wd) != 8 or not wd.isdigit():
                self._show_error("입력 오류", "날짜는 8자리 숫자로 입력해 주세요.")
                return
            entries[idx].update({
                "start": s, "end": e, "method": var_method.get(),
                "direction": var_dir.get(), "lane": var_lane.get(), "work_date": wd,
            })
            sec = route.get("detail_section_km")
            dc  = route.get("detail_container")
            if sec is not None and dc and dc.winfo_ismapped():
                self.draw_detail_schematic(route, sec)
            dlg.destroy()

        def do_delete():
            entries.pop(idx)
            sec = route.get("detail_section_km")
            dc  = route.get("detail_container")
            if sec is not None and dc and dc.winfo_ismapped():
                self.draw_detail_schematic(route, sec)
            dlg.destroy()

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.grid(row=len(rows), column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self._create_button(btn_row, text="저장", command=do_save, height=32,
                      corner_radius=8).pack(side="left", fill="x", expand=True, padx=(0, 4))
        self._create_button(btn_row, text="삭제", command=do_delete, height=32,
                      corner_radius=8, fg_color="#E53E3E", hover_color="#C53030").pack(
            side="left", fill="x", expand=True, padx=(4, 0))

        dlg.update_idletasks()
        w = max(dlg.winfo_reqwidth() + 40, 320)
        h = dlg.winfo_reqheight() + 40
        x = self.winfo_x() + (self.winfo_width() - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

    def exit_magnifier_mode(self, event=None):
        """ESC로 돋보기 모드 강제 종료 + 상세 모식도 초기화."""
        if not self.magnifier_mode:
            return
        self.magnifier_mode = False
        for route in self.routes:
            c = route.get("canvas")
            if c:
                c.configure(cursor="")
                c.bind("<Button-1>", lambda e: e.widget.focus_set())
                c.unbind("<Motion>")
                c.unbind("<Leave>")
            # 상세 구간 초기화 후 hint 텍스트 복원
            route["detail_section_km"] = None
            dc = route.get("detail_canvas")
            if dc:
                try:
                    dc.delete("all")
                    w = dc.winfo_width() or 900
                    h = dc.winfo_height() or 130
                    dc.create_text(
                        w // 2, h // 2,
                        text="10m 단위 상세보기를 원하시면 키보드 O를 누른 후 원하는 구간을 클릭하세요 (지사 보수현황 확인)",
                        fill="#AAAAAA", font=(self.font_family, 11), anchor="center", tags="hint_text"
                    )
                except Exception:
                    pass
            detail_title = route.get("detail_title_lbl")
            if detail_title:
                try:
                    detail_title.configure(text="")
                except Exception:
                    pass
            self._stop_magnifier_highlight(route)
        self._hide_magnifier_tooltip()
        self.focus_set()  # 포커스를 메인 윈도우로 복귀 → 이후 o 키 정상 인식
        # 임시 커서 파일 정리
        if self._magnifier_cur_path and os.path.exists(self._magnifier_cur_path):
            try:
                os.remove(self._magnifier_cur_path)
            except Exception:
                pass
            self._magnifier_cur_path = None

    def _on_magnifier_motion(self, event):
        """돋보기 모드에서 마우스 이동 시 툴팁 갱신."""
        if not self.magnifier_mode:
            return
        rx = event.widget.winfo_rootx() + event.x + 18
        ry = event.widget.winfo_rooty() + event.y + 18
        if self.magnifier_tooltip is None:
            self.magnifier_tooltip = tk.Toplevel(self)
            self.magnifier_tooltip.overrideredirect(True)
            self.magnifier_tooltip.attributes("-topmost", True)
            try:
                self._apply_window_icon(self.magnifier_tooltip)
            except Exception:
                pass
            tk.Label(self.magnifier_tooltip, text="상세보기(10m 단위)",
                     bg="#FFFFC0", fg="#333333", font=(self.font_family, 9),
                     relief="solid", bd=1, padx=5, pady=3).pack()
        self.magnifier_tooltip.geometry(f"+{rx}+{ry}")
        self.magnifier_tooltip.deiconify()

    def _hide_magnifier_tooltip(self, event=None):
        if self.magnifier_tooltip:
            self.magnifier_tooltip.withdraw()

    def on_magnifier_click(self, event):
        """돋보기로 모식도 클릭 → 해당 1km 구간을 10m 단위 상세 모식도로 표시."""
        event.widget.focus_set()
        if not self.magnifier_mode:
            return
        widget = event.widget
        if widget not in self.canvas_to_route_index:
            return
        route_idx = self.canvas_to_route_index[widget]
        route = self.routes[route_idx]

        mapping = route.get("_km_mapping")
        if not mapping:
            return

        cx = widget.canvasx(event.x)
        clicked_km = self._px_to_km(cx, mapping)
        if clicked_km is None:
            return

        # 1km 단위로 스냅 (노선 범위 클램프)
        section_km = int(math.floor(clicked_km))
        section_km = max(int(math.floor(mapping["r_start"])),
                         min(section_km, int(math.ceil(mapping["r_end"])) - 1))
        self.draw_detail_schematic(route, section_km)

    def _px_to_km(self, x_px: float, mapping: dict):
        """픽셀 x 좌표 → km 역산 (IC 갭 보정 반복 근사)."""
        r_start   = mapping["r_start"]
        km_marks  = mapping["km_marks"]
        ic_gap_w  = mapping["ic_gap_w"]
        px_per_km = mapping["px_per_km"]
        margin = 20

        x_raw = x_px - margin
        km_approx = r_start + x_raw / px_per_km
        for _ in range(15):
            n_gaps = sum(1 for gk in km_marks if km_approx >= gk - 1e-9)
            x_adj = x_raw - n_gaps * ic_gap_w
            km_new = r_start + x_adj / px_per_km
            if abs(km_new - km_approx) < 1e-7:
                break
            km_approx = km_new
        return km_approx

    def draw_detail_schematic(self, route: dict, section_km: int):
        """10m 단위 상세 모식도 그리기 (1km 구간)."""
        detail_canvas    = route.get("detail_canvas")
        detail_container = route.get("detail_container")
        detail_title_lbl = route.get("detail_title_lbl")
        detail_hbar      = route.get("detail_hbar")
        if not detail_canvas or not detail_container:
            return

        # 처음 호출 시 컨테이너를 탭에 붙임
        if not detail_container.winfo_ismapped():
            detail_container.pack(fill="x", pady=(4, 0))

        route["detail_section_km"] = section_km

        r_start = float(section_km)
        r_end   = float(section_km + 1)

        # 메인 캔버스 구간 하이라이트 시작
        self._start_magnifier_highlight(route, r_start, r_end)

        if detail_title_lbl:
            detail_title_lbl.configure(
                text=f"▼ 상세보기: {section_km}km ~ {section_km + 1}km  (10m 단위)",
                text_color="#FFFFFF")

        DETAIL_GRID_KM = 0.01   # 10m
        # 지사(상세) 모식도도 디스플레이 스케일에 맞춰 세로/여백을 줄여, 소형에서
        # 상세 캔버스(DETAIL_CANVAS_H)를 넘어 아래가 잘리지 않도록 한다.
        import constants as _C
        _s = float(getattr(_C, "DISPLAY_SCALE", 1.0) or 1.0)
        margin     = max(10, round(30 * _s))
        bar_h      = max(24, round(67 * _s))
        top_margin = max(14, round(32 * _s))
        dir_gap    = max(6,  round(14 * _s))
        lbl_off    = max(6,  round(12 * _s))   # km 라벨 바깥 여백
        m_off      = max(5,  round(8  * _s))   # 50m 눈금 라벨 여백
        fs_lbl     = max(6,  round(9  * _s))   # km 라벨 폰트
        fs_small   = max(6,  round(7  * _s))   # 50m/연도 라벨 폰트
        bar1_top   = top_margin
        bar2_top   = bar1_top + bar_h + dir_gap
        bar_bottom = bar2_top + bar_h

        # 창 너비에 꽉 차게 (스크롤 없음)
        detail_canvas.update_idletasks()
        canvas_w = detail_canvas.winfo_width()
        if canvas_w < 100:
            canvas_w = 900
        DETAIL_PX_PER_KM = max(100, canvas_w - margin * 2)
        total_w = canvas_w

        detail_canvas.delete("all")
        detail_canvas.config(scrollregion=(0, 0, total_w, bar_bottom + 50))

        def x_of_km_d(km_val: float) -> int:
            return int((km_val - r_start) * DETAIL_PX_PER_KM) + margin

        right = total_w - margin
        dir0  = route["directions"][0]
        lane_count = route.get("lane_count", 4)

        # ── 배경 바 ──
        for y_top in (bar1_top, bar2_top):
            detail_canvas.create_rectangle(margin, y_top, right, y_top + bar_h,
                                           fill=BACKGROUND_FILL, outline="")

        # ── 차로 구분선 ──
        for y_top in (bar1_top, bar2_top):
            for j in range(1, lane_count):
                y = y_top + bar_h * j / lane_count
                detail_canvas.create_line(margin, y, right, y, fill=LANE_BORDER)

        # ── 중분대 (gap 내부) ──
        median_h   = bar_h / max(lane_count, 1)
        draw_med_h = min(median_h, dir_gap / 2)
        for y1, y2 in [(bar1_top + bar_h, bar1_top + bar_h + draw_med_h),
                       (bar2_top - draw_med_h, bar2_top)]:
            detail_canvas.create_rectangle(margin, y1, right, y2,
                                           fill="#000000", outline="")

        # ── 10m 그리드 (50m마다 회색 실선, 나머지는 빨간 점선) ──
        x = r_start
        step_idx = 0
        while x <= r_end + 1e-9:
            xx = x_of_km_d(x)
            is_start = abs(x - r_start) < 1e-9
            is_end   = abs(x - r_end)   < 1e-9
            is_50m   = (step_idx % 5 == 0) and not is_start and not is_end
            for y_top in (bar1_top, bar2_top):
                if is_50m:
                    detail_canvas.create_line(xx, y_top, xx, y_top + bar_h,
                                              fill="#666666", width=1)
                elif not is_start and not is_end:
                    detail_canvas.create_line(xx, y_top, xx, y_top + bar_h,
                                              fill=GRID_100M, dash=(2, 4))
            if is_start or is_end:
                lbl = f"{x:.0f}k"
                detail_canvas.create_text(xx, bar1_top - lbl_off, text=lbl,
                                          fill=TEXT_COLOR, font=(self.font_family, fs_lbl, "bold"))
                detail_canvas.create_text(xx, bar2_top + bar_h + lbl_off, text=lbl,
                                          fill=TEXT_COLOR, font=(self.font_family, fs_lbl, "bold"))
            elif is_50m:
                m = step_idx * 10
                detail_canvas.create_text(xx, bar1_top - m_off, text=str(m),
                                          fill="#888888", font=(self.font_family, fs_small))
            x = round(x + DETAIL_GRID_KM, 6)
            step_idx += 1

        # ── 지사 이력 그리기 (branch_entries) ──
        route["branch_item_to_index"] = {}
        for idx, it in enumerate(route.get("branch_entries", [])):
            s, e = it.get("start", 0.0), it.get("end", 0.0)
            if e <= r_start - 1e-9 or s >= r_end + 1e-9:
                continue
            x1 = x_of_km_d(max(s, r_start))
            x2 = x_of_km_d(min(e, r_end))
            if x2 <= x1:
                x2 = x1 + 2
            style = BRANCH_METHOD_STYLES.get(it.get("method", ""), {"fill": "#808080"})
            top   = bar1_top if it.get("direction") == dir0 else bar2_top
            if it.get("lane") == "전차로":
                y1, y2 = top, top + bar_h
            else:
                try:
                    n = int(it["lane"].replace("차로", ""))
                except (ValueError, AttributeError, KeyError):
                    n = 1
                n = max(1, min(n, lane_count))
                seg_h = bar_h / lane_count
                if top == bar1_top:
                    y1 = top + bar_h - n * seg_h + 2
                    y2 = top + bar_h - (n - 1) * seg_h - 2
                else:
                    y1 = top + (n - 1) * seg_h + 2
                    y2 = top + n * seg_h - 2
            item_id = detail_canvas.create_rectangle(x1, y1, x2, y2,
                                                     fill=style["fill"], outline="",
                                                     tags=("branch_entry",))
            route["branch_item_to_index"][item_id] = idx
            # 연도 표시 (기존 모식도와 동일: '25 형식)
            wd = str(it.get("work_date") or "")
            year = wd[:4] if len(wd) >= 4 else None
            if year and x2 - x1 > 24:
                fill_hex = style.get("fill", "#808080")
                try:
                    rv, gv, bv = int(fill_hex[1:3],16), int(fill_hex[3:5],16), int(fill_hex[5:7],16)
                    tc = "#FFFFFF" if (rv*299 + gv*587 + bv*114)/1000 < 140 else "#000000"
                except Exception:
                    tc = "#FFFFFF"
                detail_label = "계획" if it.get("plan") else f"'{year[2:]}"
                detail_canvas.create_text(
                    (x1+x2)/2, (y1+y2)/2, text=detail_label,
                    fill=tc, font=(self.font_family, fs_small), anchor="center",
                    tags=("branch_label",))

        if detail_hbar:
            detail_canvas.configure(xscrollcommand=detail_hbar.set)

    # ──────────────────────────────────────────────

    def draw_schematic_for_pdf(self, canvas: tk.Canvas, route: dict, r_start: float, r_end: float, page_width: float, view_mode: str, route_start_total: float, route_end_total: float, view_hpci: bool, view_di: bool, view_aar: bool, structure_labels_to_draw: list, segment_index: int, view_rd: bool = False, view_iri: bool = False):
        """PDF 내보내기를 위해 지정된 구간의 모식도만 그리는 함수"""
        canvas.delete("all")
        span_km = max(0.1, r_end - r_start)

        selected_year = "모두" # PDF에서는 항상 모든 연도 표시

        ic_gap_w = IC_BOX_W + IC_GAP_MARGIN * 2
        ic_list = self._collect_ic_list(route, r_start, r_end)
        total_gaps_px = 0 # PDF에서는 IC 갭을 그리지 않으므로 0으로 설정

        effective_page_width = page_width * 0.95 - 40 # 좌우 여백 고려
        # 항상 7km를 기준으로 스케일을 고정하여 짧은 구간이 늘어나지 않도록 합니다.
        px_per_km_pdf = effective_page_width / 7.0

        total_w = int(span_km * px_per_km_pdf) + 80 + total_gaps_px
        total_h = 450 # PDF 출력용 고정 높이

        top_margin = 60
        bar_h = 110 # 각 방향 바 높이 (메인 화면과 동일하게 110으로 조정)
        dir_gap = 40 # PDF 출력에서는 방향 간 간격을 줄여서 한 세트를 묶어줌
        bar1_top = top_margin + 20 # 구조물명 표시 공간 확보를 위해 전체적으로 아래로 이동
        bar2_top = bar1_top + bar_h + dir_gap
        bar_bottom = bar2_top + bar_h

        canvas.config(scrollregion=(0, 0, total_w, total_h))

        km_marks = [km for km, _ in ic_list]
        def offset_for(km_value: float) -> int:
            cnt = 0
            for gk in km_marks:
                if km_value >= gk - 1e-9:
                    cnt += 1
            return cnt * ic_gap_w

        def x_of_km(km_value: float) -> int:
            return int((km_value - r_start) * px_per_km_pdf) + 20 + offset_for(km_value)

        gaps_px = [(max(20, x_of_km(km_ic) - ic_gap_w), min(total_w - 20, x_of_km(km_ic)))
                   for km_ic, _ic in ic_list]
        gaps_px = _merge_gaps(sorted(gaps_px))

        def in_gap(xpx: int) -> bool:
            for a, b in gaps_px:
                if a <= xpx <= b:
                    return True
            return False

        x = km_ceil_to_grid(r_start, GRID_KM)
        grid_positions = []
        while x <= r_end + 1e-9:
            xx = x_of_km(x)
            if not in_gap(xx):
                canvas.create_line(xx, bar1_top, xx, bar1_top + bar_h, fill=GRID_100M, dash=(2, 4))
                canvas.create_line(xx, bar2_top, xx, bar2_top + bar_h, fill=GRID_100M, dash=(2, 4))
                grid_positions.append(xx)
            x += GRID_KM

        km_positions = []
        km_tick = math.ceil(r_start)
        while km_tick <= r_end + 1e-6:
            x = x_of_km(km_tick)
            if not in_gap(x):
                canvas.create_line(x, bar1_top, x, bar1_top + bar_h, fill="#000000", width=2)
                canvas.create_line(x, bar2_top, x, bar2_top + bar_h, fill="#000000", width=2)                
                canvas.create_text(x, bar1_top - 16, text=f"{km_tick:.0f}k", fill=TEXT_COLOR, font=(self.font_family, 9))
                canvas.create_text(x, bar2_top + bar_h + 16, text=f"{km_tick:.0f}k", fill=TEXT_COLOR, font=(self.font_family, 9))
                km_positions.append(x)
            km_tick += 1.0

        lane_count = route.get("lane_count", 4)
        lane_y_positions = self._draw_bar_bg(canvas, gaps_px, total_w, bar_h, bar1_top, bar2_top, lane_count)

        # 방향 라벨 - 한글이므로 label_collector에 기록
        structure_labels_to_draw.append({
            "type": "direction",
            "text": route['directions'][0],
            "cx": 24, "cy": bar1_top - 36,
            "font_size": 10, "bold": True, "color": TEXT_COLOR,
            "anchor": "w", "segment": segment_index,
            "canvas_w": total_w, "canvas_h": total_h,
        })
        structure_labels_to_draw.append({
            "type": "direction",
            "text": route['directions'][1],
            "cx": 24, "cy": bar2_top + bar_h + 36,
            "font_size": 10, "bold": True, "color": TEXT_COLOR,
            "anchor": "w", "segment": segment_index,
            "canvas_w": total_w, "canvas_h": total_h,
        })

        # IC/JCT 그리기 (세로 쓰기, 램프 제외)
        for km_ic, ic in ic_list:
            right_edge = x_of_km(km_ic)
            left_edge = right_edge - ic_gap_w
            box_left = left_edge + IC_GAP_MARGIN
            box_right = right_edge - IC_GAP_MARGIN
            box_top = bar1_top
            box_bottom = bar2_top + bar_h

            canvas.create_rectangle(box_left, box_top, box_right, box_bottom, outline=IC_COLOR, width=3, fill="#FFFFFF")

            center_x = int((box_left + box_right) / 2)
            name = str(ic.get('name') or '')
            km_str = f"{km_ic:.1f}k"
            char_h = 16
            gap = 10
            total_text_h = (len(name) + len(km_str)) * char_h + gap
            box_h = box_bottom - box_top
            y_cursor = box_top + (box_h - total_text_h) / 2

            # IC명 (한글 포함) → label_collector에 기록, 캔버스에는 그리지 않음
            structure_labels_to_draw.append({
                "type": "ic_name",
                "text": name,
                "cx": center_x, "cy_start": y_cursor + char_h / 2,
                "char_h": char_h, "font_size": 11, "bold": True, "color": IC_COLOR,
                "segment": segment_index,
                "canvas_w": total_w, "canvas_h": total_h,
            })
            y_cursor += len(name) * char_h + gap

            # km 숫자 (ASCII) → 캔버스에 직접 그림
            for char in km_str:
                canvas.create_text(center_x, y_cursor + char_h / 2, text=char, fill=IC_COLOR, font=(self.font_family, 10, "bold"), anchor="center")
                y_cursor += char_h

        if view_mode == "기본":
            self._draw_entries(canvas, route, bar1_top, bar2_top, bar_h, lane_count, x_of_km,
                               gaps_px, "모두", clip_range=(r_start, r_end), always_label=True)

        elif view_mode == "CONDITION":
            self._draw_condition_bands(
                canvas, route, bar1_top, bar2_top, bar_h, lane_count, x_of_km,
                view_hpci, view_di, view_aar,
                view_rd=view_rd, view_iri=view_iri,
                r_start=r_start, r_end=r_end, pdf_mode=True)

        # 오버레이 (격자, 눈금선, 차로선 등)
        self._draw_overlay(canvas, grid_positions, km_positions, lane_y_positions,
                           gaps_px, total_w, bar1_top, bar2_top, bar_h, km_line_color="#000000")

        def draw_bar_outline(y_top: int):
            y_bottom = y_top + bar_h; seg_start = 20
            for a, b in gaps_px:
                if a > seg_start:
                    canvas.create_line(seg_start, y_top, a, y_top, fill="#718096")
                    canvas.create_line(seg_start, y_bottom, a, y_bottom, fill="#718096")
                seg_start = b
            if seg_start < total_w - 20:
                canvas.create_line(seg_start, y_top, total_w - 20, y_top, fill="#718096")
                canvas.create_line(seg_start, y_bottom, total_w - 20, y_bottom, fill="#718096")
        draw_bar_outline(bar1_top)
        draw_bar_outline(bar2_top)

        # 구조물 그리기 (PDF용) - 한글 이름은 label_collector에 수집 후 reportlab으로 출력
        self._draw_structures(canvas, route, r_start, r_end, bar1_top, bar2_top, bar_h,
                              x_of_km, init_font_size=8, min_font_size=4, min_width=5,
                              label_collector=structure_labels_to_draw,
                              canvas_w=total_w, canvas_h=total_h, segment_index=segment_index)
            
        # 노선 범위를 벗어나는 구간에 빗금 표시
        if r_start < route_start_total:
            cut = min(r_end, route_start_total)
            for bt in (bar1_top, bar2_top):
                _draw_hatch_rect(canvas, x_of_km(r_start), bt, x_of_km(cut), bt + bar_h)
        if r_end > route_end_total:
            cut = max(r_start, route_end_total)
            for bt in (bar1_top, bar2_top):
                _draw_hatch_rect(canvas, x_of_km(cut), bt, x_of_km(r_end), bt + bar_h)

        # 전체 노선의 시작점에 'Start' 라벨 표시
        if r_start <= route_start_total < r_end:
            x_start_label = x_of_km(route_start_total)
            canvas.create_text(x_start_label, bar_bottom + 20, anchor="w", text=f"Start {fmt_km(route_start_total)}k", fill=TEXT_COLOR, font=(self.font_family, 9))
        # 전체 노선의 끝점에 'End' 라벨 표시
        if r_start < route_end_total <= r_end:
            x_end_label = x_of_km(route_end_total)
            canvas.create_text(x_end_label, bar_bottom + 20, anchor="e", text=f"End {fmt_km(route_end_total)}k", fill=TEXT_COLOR, font=(self.font_family, 9))
        canvas.tag_raise("entry_label")
        canvas.tag_raise("hpci_label")
        canvas.tag_raise("di_label")
        canvas.tag_raise("aar_label")
        # 100m 점선 그리드: 모든 그리기 작업이 끝난 후 마지막에 직접 그려 최상단 보장
        try:
            for xx in grid_positions:
                canvas.create_line(xx, bar1_top, xx, bar1_top + bar_h, fill=GRID_100M, dash=(2, 4))
                canvas.create_line(xx, bar2_top, xx, bar2_top + bar_h, fill=GRID_100M, dash=(2, 4))
        except Exception:
            pass

    # ---------- 인터랙션: 더블클릭 상세보기/수정 ----------_
    def on_canvas_double_click(self, event):
        try:
            # 어떤 탭/캔버스인지 파악
            widget = event.widget
            if not isinstance(widget, tk.Canvas):
                return
            if widget not in self.canvas_to_route_index:
                return
            route_idx = self.canvas_to_route_index[widget]
            self.current_route_index = route_idx
            route = self.routes[route_idx]

            # 스크롤 보정 좌표
            cx = widget.canvasx(event.x)
            cy = widget.canvasy(event.y)
            # 포인터 아래 겹치는 'entry' 항목 전부 수집 (본선 + 램프)
            items = widget.find_overlapping(cx, cy, cx, cy)
            entry_items = [it for it in items if it in route["canvas_item_to_index"]]

            # 중복 인덱스 제거, 그리기 순서상 위에 있는 항목이 뒤에 있으므로 뒤에서부터 정렬
            seen = set()
            overlapped_indices = []
            for it in reversed(entry_items):
                idx = route["canvas_item_to_index"][it]
                if idx not in seen:
                    seen.add(idx)
                    overlapped_indices.append(idx)

            if not overlapped_indices:
                return
            if len(overlapped_indices) == 1:
                self.open_entry_dialog(overlapped_indices[0])
            else:
                # 여러 개면 각각 개별 창을 비모달로 띄우되, 화면에서 서로 겹치지 않도록 배치
                base_x = getattr(event, 'x_root', widget.winfo_rootx() + event.x)
                base_y = getattr(event, 'y_root', widget.winfo_rooty() + event.y)
                existing_rects = []  # (x1,y1,x2,y2)
                created = []
                for k, idx in enumerate(overlapped_indices):
                    dlg = self.open_entry_dialog(idx, modal=False)
                    try:
                        dlg.update_idletasks()
                        w = dlg.winfo_width() or dlg.winfo_reqwidth() or 320
                        h = dlg.winfo_height() or dlg.winfo_reqheight() or 260
                        sw = dlg.winfo_screenwidth()
                        sh = dlg.winfo_screenheight()
                    except Exception:
                        w, h = 320, 260
                        sw, sh = 1920, 1080

                    # 후보 위치를 그리드 탐색해 겹치지 않게 배치
                    gap = 12
                    placed = False
                    for row in range(0, 8):
                        for col in range(0, 6):
                            x = base_x + col * (w + gap)
                            y = base_y + row * (h + gap)
                            # 화면 안으로 보정
                            x = max(0, min(x, sw - w - 8))
                            y = max(0, min(y, sh - h - 8))
                            # 기존 사각형들과 겹침 검사
                            overlap = False
                            for (rx1, ry1, rx2, ry2) in existing_rects:
                                if not (x + w <= rx1 or rx2 <= x or y + h <= ry1 or ry2 <= y):
                                    overlap = True
                                    break
                            if not overlap:
                                dlg.geometry(f"+{int(x)}+{int(y)}")
                                existing_rects.append((x, y, x + w, y + h))
                                placed = True
                                break
                        if placed:
                            break
                    if not placed:
                        # 실패 시 대각선으로 조금씩 벌리며 배치
                        x = max(0, min(base_x + k * (gap + 20), sw - w - 8))
                        y = max(0, min(base_y + k * (gap + 20), sh - h - 8))
                        dlg.geometry(f"+{int(x)}+{int(y)}")
                        existing_rects.append((x, y, x + w, y + h))
                    created.append(dlg)
        except Exception:
            pass

    def on_tree_double_click(self, event):
        try:
            # 어떤 탭인지
            widget = event.widget
            idx_tab = self.current_route_index
            if not (0 <= idx_tab < len(self.routes)):
                return
            route = self.routes[idx_tab]
            tree = route.get("list_tree")
            if tree is None or widget is not tree:
                return
            sel = tree.selection()
            if not sel:
                return
            item_id = sel[0]
            if item_id in route.get("tree_item_to_index", {}):
                entry_idx = route["tree_item_to_index"][item_id]
                self.open_entry_dialog(entry_idx)
        except Exception:
            pass

    def on_open_detail_table(self):
        # 현재 탭의 이력을 표로 표시하는 별도 창
        if not self.routes:
            self._show_info("안내", "표시할 노선이 없습니다.")
            return

        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.transient(self)
        dlg.geometry("1200x500")
        sort_state = {"col": "work_date", "reverse": True}
        
        # 상단: 필터 바(연도는 좌측 메인과 연동)
        filter_bar = ctk.CTkFrame(dlg, fg_color="transparent")
        filter_bar.pack(fill="x", padx=10, pady=10)
        
        all_entries = []
        for r in self.routes:
            for entry in r.get("entries", []):
                entry['route_name'] = r['name']
                all_entries.append(entry)
        
        filter_vars = {}
        var_search = tk.StringVar()
        ctk.CTkLabel(filter_bar, text="검색").pack(side="left", padx=(0, 4))
        ent_search = ctk.CTkEntry(filter_bar, textvariable=var_search, width=180,
                                  placeholder_text="노선, 공법, 방향, 날짜")
        ent_search.pack(side="left", padx=(0, 10))
        
        def create_multi_select_filter(label, options):
            ctk.CTkLabel(filter_bar, text=label).pack(side="left", padx=(14, 0))
            menubutton = self._create_dropdown_button(filter_bar, "모두", width=140, height=30)
            menu = self._create_styled_menu(menubutton)
            menubutton.configure(command=lambda: self._show_menu_below_widget(menubutton, menu))
            
            vars_dict = {opt: tk.BooleanVar(value=True) for opt in options}
            filter_vars[label] = (vars_dict, menubutton)
            
            def update_selection_text():
                selected = [opt for opt, var in vars_dict.items() if var.get()]
                if len(selected) == len(options):
                    label_text = "모두"
                elif len(selected) > 1:
                    label_text = f"{len(selected)}개 선택됨"
                elif len(selected) == 1:
                    label_text = selected[0]
                else:
                    label_text = "선택 없음"
                menubutton.configure(text=f"{label_text}  ▾")
                populate_table()
                menubutton.after(10, lambda: self._show_menu_below_widget(menubutton, menu))

            
            def toggle_all():
                all_selected = all(v.get() for v in vars_dict.values())
                new_state = not all_selected
                for var in vars_dict.values():
                    var.set(new_state)
                update_selection_text()
            
            menu.add_command(label="전체 선택/해제", command=toggle_all)
            menu.add_separator()
            
            for option in options:
                menu.add_checkbutton(label=option, variable=vars_dict[option], onvalue=True, offvalue=False, command=update_selection_text)
            
            menubutton.pack(side="left", padx=6)
        
        # 필터 생성
        route_names = sorted(list(set(r['name'] for r in self.routes)))
        create_multi_select_filter("노선", route_names)
        
        years = sorted(list(set(
            (wd[:4]) for it in all_entries if (wd := str(it.get("work_date") or "")) and len(wd) >= 4
        )))
        create_multi_select_filter("연도", years)
        
        dir_values = sorted(list(set(
            d for it in all_entries if (d := str(it.get("direction") or "").strip())
        )))
        create_multi_select_filter("방향", dir_values)
        
        lane_values = sorted(list(set(
            ln for it in all_entries if (ln := str(it.get("lane") or "").strip())
        )))
        create_multi_select_filter("차로", lane_values)
        
        method_values = sorted(list(set(
            m for it in all_entries if (m := str(it.get("method") or "").strip())
        )))
        create_multi_select_filter("공법", method_values)
        
        # 열 구성: 노선명, 이정, 시공연장(m) 등
        cols = ("route_name", "range", "len_m", "direction", "lane", "method", "work_date")
        tree = ttk.Treeview(dlg, columns=cols, show="headings")
        tree.heading("route_name", text="노선명")
        tree.heading("range", text="이정(km)")
        tree.heading("len_m", text="시공연장(m)")
        tree.heading("direction", text="방향")
        tree.heading("lane", text="차로")
        tree.heading("method", text="공법")
        tree.heading("work_date", text="시공일자")
        
        tree.column("route_name", width=120, anchor="center", stretch=False)
        tree.column("range", width=180, anchor="center", stretch=False)
        tree.column("len_m", width=120, anchor="center", stretch=False)
        tree.column("direction", width=120, anchor="center", stretch=False)
        tree.column("lane", width=100, anchor="center", stretch=False)
        tree.column("method", width=160, anchor="center", stretch=False)
        tree.column("work_date", width=120, anchor="center", stretch=False)
        
        # 가독성을 위한 줄무늬 배경
        tree.tag_configure("odd", background="#F7FAFC")
        tree.tag_configure("even", background="#FFFFFF")
        tree.tag_configure("total", background="#E2E8F0")
        
        vbar = ctk.CTkScrollbar(dlg, orientation="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vbar.set)
        tree.pack(side="left", fill="both", expand=True)
        vbar.pack(side="left", fill="y")
        
        # Treeview item ID와 원본 entry 딕셔너리를 매핑하기 위한 딕셔너리
        iid_to_entry = {}

        def get_sort_key(row_dict, col_name):
            if col_name == "route_name":
                return str(row_dict.get("route_name", ""))
            if col_name == "range":
                return float(row_dict.get("start", 0.0))
            if col_name == "len_m":
                return max(0.0, (float(row_dict.get("end", 0)) - float(row_dict.get("start", 0))) * 1000.0)
            if col_name == "direction":
                return str(row_dict.get("direction", ""))
            if col_name == "lane":
                lane_val = str(row_dict.get("lane", ""))
                if lane_val == "전차로":
                    return 0
                try:
                    return int(lane_val.replace("차로", ""))
                except Exception:
                    return 999
            if col_name == "method":
                return str(row_dict.get("method", ""))
            if col_name == "work_date":
                return str(row_dict.get("work_date", ""))
            return ""

        def on_sort(col_name):
            if sort_state["col"] == col_name:
                sort_state["reverse"] = not sort_state["reverse"]
            else:
                sort_state["col"] = col_name
                sort_state["reverse"] = False
            populate_table()

        tree.heading("route_name", command=lambda: on_sort("route_name"))
        tree.heading("range", command=lambda: on_sort("range"))
        tree.heading("len_m", command=lambda: on_sort("len_m"))
        tree.heading("direction", command=lambda: on_sort("direction"))
        tree.heading("lane", command=lambda: on_sort("lane"))
        tree.heading("method", command=lambda: on_sort("method"))
        tree.heading("work_date", command=lambda: on_sort("work_date"))
        
        # 테이블 채우기 함수
        def populate_table(*_):
            # 초기화
            for item in tree.get_children():
                tree.delete(item)
            
            iid_to_entry.clear()
            
            # 선택된 필터 값 가져오기
            selected_filters = {}
            for label, (vars_dict, _) in filter_vars.items():
                selected_filters[label] = {opt for opt, var in vars_dict.items() if var.get()}
            
            # 노선 필터에 따라 노선 자체의 총 연장 계산 (다른 필터와 무관)
            total_route_km = 0.0
            selected_route_names = selected_filters.get("노선", set())
            for r in self.routes:
                if r.get('name') in selected_route_names:
                    total_route_km += (r.get('end_km', 0.0) - r.get('start_km', 0.0))
            total_route_km_str = f"{total_route_km:.2f}km"

            # 필터링된 행과 총연장 계산
            rows = []
            total_m = 0
            search_kwd = var_search.get().strip().lower()
            
            for it in all_entries:
                # 각 필터 조건 확인
                if selected_filters["노선"] and it.get('route_name') not in selected_filters["노선"]:
                    continue
                
                wd = str(it.get("work_date") or "")
                year = wd[:4] if len(wd) >= 4 else None
                if selected_filters["연도"] and year not in selected_filters["연도"]:
                    continue
                
                if selected_filters["방향"] and (it.get("direction") or "").strip() not in selected_filters["방향"]:
                    continue
                
                if selected_filters["차로"] and (it.get("lane") or "").strip() not in selected_filters["차로"]:
                    continue
                
                if selected_filters["공법"] and (it.get("method") or "").strip() not in selected_filters["공법"]:
                    continue
                
                rng = f"{it['start']:.3f} ~ {it['end']:.3f}"
                length_m = max(0, (float(it.get('end', 0)) - float(it.get('start', 0))) * 1000.0)
                lm_int = int(round(length_m))
                total_m += lm_int

                # 시공일자 표시는 하이픈 포함(입력값은 그대로 유지)
                if len(wd) == 8 and wd.isdigit():
                    wd_disp = f"{wd[:4]}-{wd[4:6]}-{wd[6:8]}"
                else:
                    wd_disp = wd

                search_blob = " ".join([
                    str(it.get('route_name', '')),
                    rng,
                    str(it.get("direction", "")),
                    str(it.get("lane", "")),
                    str(it.get("method", "")),
                    wd_disp
                ]).lower()
                if search_kwd and search_kwd not in search_blob:
                    continue
                
                rows.append((it.get('route_name', ''), rng, lm_int, it.get("direction", ""), it.get("lane", ""), it.get("method", ""), wd_disp, it))

            rows.sort(key=lambda row: get_sort_key(row[7], sort_state["col"]), reverse=sort_state["reverse"])
            
            # 총계 행 추가 (맨 위): 이정은 비우고, 시공연장 합계만 표시
            total_vals_disp = ("계", total_route_km_str, f"{total_m:,}", "", "", "", "")
            tree.insert("", 0, values=total_vals_disp, tags=("total",))
            
            # 데이터 행 추가
            row_idx = 0
            for vals in rows:
                tag = "odd" if (row_idx % 2 == 0) else "even"
                route_name, rng, lm_int, d, ln, mth, wd_disp, original_entry = vals
                vals_disp = (route_name, rng, f"{lm_int:,}", d, ln, mth, wd_disp)
                iid = tree.insert("", "end", values=vals_disp, tags=(tag,))
                # iid와 원본 entry를 매핑
                iid_to_entry[iid] = original_entry
                row_idx += 1
            if hasattr(dlg, "_detail_summary_lbl"):
                data_rows = max(0, len(tree.get_children()) - 1)
                sort_dir = "내림차순" if sort_state["reverse"] else "오름차순"
                dlg._detail_summary_lbl.configure(
                    text=f"조회 {data_rows}건 | 정렬: {sort_state['col']} {sort_dir}"
                )
        
        # 최초 로드
        populate_table()

        action_bar = ctk.CTkFrame(dlg, fg_color="transparent")
        action_bar.pack(fill="x", padx=10, pady=(0, 8))

        summary_lbl = ctk.CTkLabel(action_bar, text="", text_color="gray70")
        summary_lbl.pack(side="left")
        dlg._detail_summary_lbl = summary_lbl

        def refresh_summary():
            data_rows = max(0, len(tree.get_children()) - 1)
            sort_dir = "내림차순" if sort_state["reverse"] else "오름차순"
            summary_lbl.configure(text=f"조회 {data_rows}건 | 정렬: {sort_state['col']} {sort_dir}")

        def open_selected_entry():
            sel = tree.selection()
            if not sel:
                return
            iid = sel[0]
            original_entry = iid_to_entry.get(iid)
            if not isinstance(original_entry, dict):
                return

            route_name = original_entry.get('route_name')
            target_route = None
            target_route_idx = -1
            for i, r in enumerate(self.routes):
                if r['name'] == route_name:
                    target_route = r
                    target_route_idx = i
                    break

            if target_route and original_entry in target_route.get('entries', []):
                entry_idx = target_route['entries'].index(original_entry)
                self.open_entry_dialog(entry_idx, route_idx_override=target_route_idx)

        self._create_button(action_bar, text="선택 수정", command=open_selected_entry,
                      width=90, height=28).pack(side="right")
        self._create_button(action_bar, text="새로고침", command=populate_table,
                      width=90, height=28, fg_color="#4A5568").pack(side="right", padx=(0, 6))
        ent_search.bind("<Return>", populate_table)
        try:
            var_search.trace_add("write", lambda *_: populate_table())
        except Exception:
            pass
        refresh_summary()

        def on_row_double(_):
            open_selected_entry()

        tree.bind("<Double-1>", on_row_double)
        tree.bind("<Return>", on_row_double)

    def open_entry_dialog(self, idx: int, modal: bool = True, pos=None, route_idx_override=None):
        # 현재 노선의 항목으로 접근
        if route_idx_override is not None:
            current_route_idx = route_idx_override
        else:
            current_route_idx = self.current_route_index

        route = self.routes[current_route_idx]
        it = route["entries"][idx]

        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.transient(self)
        if modal:
            try:
                dlg.grab_set()
            except Exception:
                pass

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=10)

        # Vars (1 m 단위 표시: 소수점 셋째 자리까지)
        var_start = tk.StringVar(value=f"{it['start']:.3f}")
        var_end = tk.StringVar(value=f"{it['end']:.3f}")
        var_method = tk.StringVar(value=it['method'])
        var_dir = tk.StringVar(value=it['direction'])
        var_lane = tk.StringVar(value=it['lane'])
        var_wd = tk.StringVar(value=it.get('work_date', ''))
        # 길이(m) 표시용
        var_len = tk.StringVar(value="")

        for col in range(4):
            body.columnconfigure(col, weight=1)

        # Row 0: 구간
        ctk.CTkLabel(body, text="시작(km)").grid(row=0, column=0, sticky="w", padx=4, pady=3)
        ent_start = ctk.CTkEntry(body, textvariable=var_start, width=80)
        ent_start.grid(row=0, column=1, sticky="w", padx=4, pady=3)

        ctk.CTkLabel(body, text="끝(km)").grid(row=0, column=2, sticky="w", padx=4, pady=3)
        ent_end = ctk.CTkEntry(body, textvariable=var_end, width=80)
        ent_end.grid(row=0, column=3, sticky="w", padx=4, pady=3)

        # Row 1: 공법/방향
        ctk.CTkLabel(body, text="공법").grid(row=1, column=0, sticky="w", padx=4, pady=3)
        cb_method = self._create_styled_combobox(body, variable=var_method, values=list(METHOD_STYLES.keys()), width=120, state="readonly")
        cb_method.grid(row=1, column=1, sticky="w", padx=4, pady=3)

        ctk.CTkLabel(body, text="방향").grid(row=1, column=2, sticky="w", padx=4, pady=3)
        # 현재 노선의 방향 목록을 사용 (예: 청주/영덕 등 커스텀 방향 지원)
        cb_dir = self._create_styled_combobox(body, variable=var_dir, values=route.get("directions", DIRECTIONS), width=120, state="readonly")
        cb_dir.grid(row=1, column=3, sticky="w", padx=4, pady=3)

        # Row 2: 차로/시공날짜
        ctk.CTkLabel(body, text="차로").grid(row=2, column=0, sticky="w", padx=4, pady=3)
        lane_count = route.get("lane_count", 4)
        lane_values = ["전차로"] + [f"{i}차로" for i in range(1, lane_count + 1)]
        cb_lane = self._create_styled_combobox(body, variable=var_lane, values=lane_values, width=120, state="readonly")
        cb_lane.grid(row=2, column=1, sticky="w", padx=4, pady=3)

        ctk.CTkLabel(body, text="시공날짜").grid(row=2, column=2, sticky="w", padx=4, pady=3)
        ent_wd = ctk.CTkEntry(body, textvariable=var_wd, width=100)
        ent_wd.grid(row=2, column=3, sticky="w", padx=4, pady=3)

        # Row 3: 길이(m) 표시 (예: 36.00~36.20 : 200m)
        ctk.CTkLabel(body, text="길이").grid(row=3, column=0, sticky="w", padx=4, pady=3)
        lbl_len = ctk.CTkLabel(body, textvariable=var_len)
        lbl_len.grid(row=3, column=1, columnspan=3, sticky="w", padx=4, pady=3)

        # 길이 계산 함수 및 바인딩
        def recalc_len(*_):
            try:
                s_val = float(var_start.get().strip())
                e_val = float(var_end.get().strip())
                length_m = max(0, (e_val - s_val) * 1000.0)
                length_str = f"{s_val:.3f}~{e_val:.3f} : {int(round(length_m))}m"
            except Exception:
                length_str = "-"
            var_len.set(length_str)
        recalc_len()
        try:
            var_start.trace_add('write', recalc_len)
            var_end.trace_add('write', recalc_len)
        except Exception:
            pass

        # 초기에는 편집 불가 상태
        def set_editable(enabled: bool):
            state_norm = "normal" if enabled else "readonly"
            ent_start.configure(state="normal" if enabled else "disabled")
            ent_end.configure(state="normal" if enabled else "disabled")
            cb_method.configure(state=state_norm)
            cb_dir.configure(state=state_norm)
            cb_lane.configure(state=state_norm)
            ent_wd.configure(state="normal" if enabled else "disabled")

        set_editable(False)

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(fill="x", padx=10, pady=10)

        def on_save():
            try:
                s = float(var_start.get().strip())
                e = float(var_end.get().strip())
            except Exception:
                self._show_error("입력 오류", "시작/끝 이정을 숫자로 입력해 주세요. 예: 36.00, 36.20")
                return

            if e <= s:
                self._show_error("입력 오류", "끝 이정은 시작 이정보다 커야 합니다.")
                return

            r_start = float(self.route_start_km.get())
            r_end = float(self.route_end_km.get())
            if not (r_start <= s <= r_end and r_start <= e <= r_end):
                self._show_error("범위 오류", f"이정은 노선 구간 {fmt_km(r_start)} ~ {fmt_km(r_end)} km 안이어야 합니다.")
                return

            # 날짜 검증
            wd = var_wd.get().strip()
            if len(wd) != 8 or not wd.isdigit():
                self._show_error("입력 오류", "8글자로 작성해주세요 (예:20250723)")
                return

            # 스냅(1 m 단위)
            s_snapped = clamp(km_floor_to_grid(s, ENTRY_GRID_KM), r_start, r_end)
            e_snapped = clamp(km_ceil_to_grid(e, ENTRY_GRID_KM), r_start, r_end)

            new_dir = var_dir.get()
            new_lane = var_lane.get()
            if new_lane != "전차로":
                try:
                    lane_num = int(new_lane.replace("차로", ""))
                    if lane_num > route.get("lane_count", 4):
                        self._show_error("입력 오류", f"이 노선은 {route.get('lane_count', 4)}차로까지 설정되어 있습니다.")
                        return
                except ValueError: pass
            # 중복 허용: 자기 자신 외 구간과 겹쳐도 저장 가능

            # 저장
            it.update({
                "start": s_snapped,
                "end": e_snapped,
                "method": var_method.get(),
                "direction": new_dir,
                "lane": new_lane,
                "work_date": wd,
            })
            self._mark_route_data_changed(route)
            # 현재 활성화된 탭이 수정된 노선이 아닐 수 있으므로, 해당 노선을 활성화
            if self.current_route_index != current_route_idx:
                self.notebook.set(self.routes[current_route_idx]['name'])

            self.update_year_filter_values()
            self.draw_schematic()
            dlg.destroy()

        edit_mode = {"enabled": False}

        def on_edit_toggle():
            if not edit_mode["enabled"]:
                edit_mode["enabled"] = True
                set_editable(True)
                btn_edit.configure(text="저장", command=on_save)
                btn_close.configure(text="취소", command=dlg.destroy)
            else:
                on_save()

        def on_delete():
            try:
                if self._ask_yes_no("삭제 확인", "이 이력을 삭제하시겠습니까?"):
                    try:
                        del route["entries"][idx]
                    except Exception:
                        return
                    self._mark_route_data_changed(route)
                    self.update_year_filter_values()
                    # 현재 활성화된 탭이 수정된 노선이 아닐 수 있으므로, 해당 노선을 활성화
                    if self.current_route_index != current_route_idx:
                        self.notebook.set(self.routes[current_route_idx]['name'])
                    self.draw_schematic()
                    dlg.destroy()
            except Exception:
                pass

        btn_edit = self._create_button(btns, text="수정", command=on_edit_toggle, width=80)
        btn_edit.pack(side="left")
        btn_delete = self._create_button(btns, text="삭제", command=on_delete, width=80, fg_color="#C53030", hover_color="#9B2C2C")
        btn_delete.pack(side="left", padx=(8, 0))
        btn_close = self._create_close_button(btns, dlg.destroy, width=80)
        btn_close.pack(side="right")

        # 위치 지정(겹침 방지용)
        try:
            if pos and isinstance(pos, (tuple, list)) and len(pos) == 2:
                dlg.geometry(f"+{int(pos[0])}+{int(pos[1])}")
        except Exception:
            pass

        return dlg



    def on_mouse_wheel_horizontal(self, event):
        try:
            widget = event.widget
            if not isinstance(widget, tk.Canvas):
                return
            # Windows/Mac: event.delta, Linux: event.num (4=up, 5=down)
            delta_units = 0
            if hasattr(event, "delta") and event.delta != 0:
                # 양수: 휠 업 → 좌로, 음수: 휠 다운 → 우로
                steps = int(abs(event.delta) / 120) if abs(event.delta) >= 120 else 1
                direction = -1 if event.delta > 0 else 1
                delta_units = direction * steps * 3
            elif hasattr(event, "num"):
                # X11(Button-4/5)
                direction = -1 if event.num == 4 else 1
                delta_units = direction * 3

            if delta_units != 0:
                widget.xview_scroll(delta_units, "units")
                return "break"
        except Exception:
            pass

