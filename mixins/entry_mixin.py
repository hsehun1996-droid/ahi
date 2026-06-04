# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - 이력 입력 모드, 이력 추가
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


class EntryMixin:
    def generate_directions_from_name(self, route_name: str):
        # 이름이 "A B선" 꼴이면 A방향/B방향으로 유추. 그 외에는 기본값 사용
        try:
            name = route_name.strip()
            if name.endswith("선"):
                core = name[:-1]
            else:
                core = name
            # 한글 방향명 분리 추정: 예) 서산영덕 → 서산/영덕
            # 간단 규칙: 가운데에서 두 토큰으로 분리 시도 (초성/중성 구분 어려워 휴리스틱)
            # 공백으로 구분되어 있으면 공백 기준
            parts = core.split()
            if len(parts) == 2:
                a, b = parts
                return [f"{a}방향", f"{b}방향"]
            # 공백이 없다면 중간 분리 시도
            mid = len(core) // 2
            a, b = core[:mid], core[mid:]
            if a and b:
                return [f"{a}방향", f"{b}방향"]
        except Exception:
            pass
        return DIRECTIONS[:]


    # ---------- 이력입력 모드 토글 (본부 ↔ 지사) ----------
    def toggle_entry_mode(self):
        """이력입력 카드를 본부 ↔ 지사 모드로 전환."""
        if self.entry_mode == "본부":
            self.entry_mode = "지사"
            if self.method_cb:
                self.method_cb.configure(values=list(BRANCH_METHOD_STYLES.keys()))
                branch_method = self.last_entry_by_mode.get("지사", {}).get("method")
                branch_values = list(BRANCH_METHOD_STYLES.keys())
                self.input_method.set(branch_method if branch_method in branch_values else branch_values[0])
        else:
            self.entry_mode = "본부"
            if self.method_cb:
                self.method_cb.configure(values=list(METHOD_STYLES.keys()))
                main_method = self.last_entry_by_mode.get("본부", {}).get("method")
                main_values = list(METHOD_STYLES.keys())
                self.input_method.set(main_method if main_method in main_values else main_values[0])
        self._apply_recent_entry_defaults()
        self._apply_entry_mode_ui()

    def _apply_entry_mode_ui(self):
        """현재 입력 모드에 맞춰 카드 헤더의 색상과 안내문을 갱신합니다."""
        if self.entry_mode == "지사":
            title = "이력입력(지사)"
            toggle_text = "본부"
            badge_text = "지사 모드"
            badge_fg = "#FFE9C7"
            badge_text_color = "#8A4B08"
            desc = "지사 시행 이력을 등록합니다. 별도 관리되며 상세 모식도에서 함께 검토됩니다."
            desc_color = "#A26119"
            card_border = "#F0C27B"
            divider_color = "#F0C27B"
            toggle_fg = "#FFE9C7"
            toggle_hover = "#F8D9A3"
            toggle_text_color = "#8A4B08"
        else:
            title = "이력입력(본부)"
            toggle_text = "지사"
            badge_text = "본부 모드"
            badge_fg = "#DCEAFE"
            badge_text_color = PRIMARY_BLUE
            desc = "본부 확정 이력을 등록합니다. 전체 이력, 우선순위, 하자분석에 직접 반영됩니다."
            desc_color = "#3B6E9E"
            card_border = "#BFD5EE"
            divider_color = "#BFD5EE"
            toggle_fg = "#DCEAFE"
            toggle_hover = "#C9DFF8"
            toggle_text_color = PRIMARY_BLUE

        if self.form_title_lbl:
            self.form_title_lbl.configure(text=title)
        if self.form_toggle_btn:
            self.form_toggle_btn.configure(text=toggle_text, fg_color=toggle_fg, hover_color=toggle_hover, text_color=toggle_text_color)
        if self.form_mode_badge:
            self.form_mode_badge.configure(text=badge_text, fg_color=badge_fg, text_color=badge_text_color)
        if self.form_mode_desc_lbl:
            self.form_mode_desc_lbl.configure(text=desc, text_color=desc_color)
        if self.form_card:
            self.form_card.configure(border_color=card_border)
        if self.form_header_divider:
            self.form_header_divider.configure(fg_color=divider_color)

    def _remember_last_entry_inputs(self, method: str, direction: str, lane: str, work_date: str):
        self.last_entry_by_mode[self.entry_mode] = {
            "method": method,
            "direction": direction,
            "lane": lane,
            "work_date": work_date,
        }

    def _apply_recent_entry_defaults(self):
        """현재 모드의 직전 입력값을 폼에 반영합니다."""
        recent = self.last_entry_by_mode.get(self.entry_mode, {})
        direction = recent.get("direction")
        lane = recent.get("lane")
        work_date = recent.get("work_date")
        if direction:
            self.input_direction.set(direction)
        if lane:
            self.input_lane.set(lane)
        if work_date:
            self.input_work_date.set(work_date)

    # ---------- 동작 ----------
    def on_add_entry(self):
        try:
            s = float(str(self.input_start_km.get()).strip())
            e = float(str(self.input_end_km.get()).strip())
        except ValueError:
            self._show_error("입력 오류", "시작/끝 이정을 숫자로 입력해 주세요. 예: 36.00, 36.20")
            return

        if e <= s:
            self._show_error("입력 오류", "끝 이정은 시작 이정보다 커야 합니다.")
            return

        # 현재 노선 기준
        if self.current_route_index < 0:
            self._show_error("노선 없음", "먼저 노선을 추가해 주세요.")
            return
        route = self.routes[self.current_route_index]
        r_start = float(route["start_km"])
        r_end = float(route["end_km"])
        # 본선일 때만 노선 범위 검증. 램프는 램프 길이 내에서 검증
        lane_count = route.get("lane_count", 4)
        target = (self.input_target.get() or "본선").strip()
        if target == "본선":
            if not (r_start <= s <= r_end and r_start <= e <= r_end):
                self._show_error("범위 오류", f"이정은 노선 구간 {fmt_km(r_start)} ~ {fmt_km(r_end)} km 안이어야 합니다.")
                return

        # 이력 스냅(10 m 단위)
        s_snapped = clamp(km_floor_to_grid(s, ENTRY_GRID_KM), r_start, r_end)
        e_snapped = clamp(km_ceil_to_grid(e, ENTRY_GRID_KM), r_start, r_end)

        method = self.input_method.get()
        direction = self.input_direction.get()
        lane = self.input_lane.get()
        if lane != "전차로":
            try:
                lane_num = int(lane.replace("차로", ""))
                if lane_num > lane_count:
                    self._show_error("입력 오류", f"이 노선은 {lane_count}차로까지 설정되어 있습니다.")
                    return
            except ValueError:
                pass

        # 시공날짜 검증 (8자리 숫자 + 실제 유효한 날짜 확인)
        work_date = str(self.input_work_date.get()).strip()
        if len(work_date) != 8 or not work_date.isdigit():
            self._show_error("입력 오류", "8글자로 작성해주세요 (예:20250723)")
            return
        try:
            datetime.strptime(work_date, "%Y%m%d")
        except ValueError:
            self._show_error("입력 오류", f"'{work_date}'는 유효하지 않은 날짜입니다 (예:20250723)")
            return

        # 중복 허용: 동일 구간/차로/방향이 겹쳐도 추가 가능

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._remember_last_entry_inputs(method, direction, lane, work_date)

        def _prepare_next_input_defaults():
            try:
                next_start = e_snapped
                default_len = max(0.2, round(e_snapped - s_snapped, 3))
                next_end = clamp(round(next_start + default_len, 3), r_start, r_end)
                if next_end <= next_start:
                    next_end = clamp(round(next_start + 0.2, 3), r_start, r_end)
                self.input_start_km.set(f"{next_start:.3f}")
                self.input_end_km.set(f"{next_end:.3f}")
            except Exception:
                pass

        # 분기: 본선 vs 램프
        if target == "본선":
            if self.entry_mode == "지사":
                # 지사 이력 → branch_entries (돋보기 상세 모식도에만 반영)
                route.setdefault("branch_entries", []).append({
                    "start": s_snapped, "end": e_snapped, "method": method,
                    "ts": ts, "direction": direction, "lane": lane, "work_date": work_date
                })
                self._mark_route_data_changed(route)
                self.update_year_filter_values()
                sec = route.get("detail_section_km")
                dc  = route.get("detail_container")
                if sec is not None and dc and dc.winfo_ismapped():
                    self.draw_detail_schematic(route, sec)
                _prepare_next_input_defaults()
            else:
                # 본부 이력 → entries (메인 모식도 반영)
                route["entries"].append({
                    "start": s_snapped, "end": e_snapped, "method": method,
                    "ts": ts, "direction": direction, "lane": lane, "work_date": work_date
                })
                self._mark_route_data_changed(route)
                self.update_year_filter_values()
                self.draw_schematic()
                _prepare_next_input_defaults()
            return

        # 램프 선택 필요
        ic_label = (self.input_ic.get() or "").strip()
        ramp_name = (self.input_ramp.get() or "").strip()
        if not ic_label or not ramp_name:
            self._show_error("입력 오류", "IC/JCT와 램프를 선택해 주세요.")
            return
        # 대상 IC 찾기
        ic_idx = -1
        try:
            ic_idx = [f"{ic.get('route','')}: {ic.get('name','')} ({float(ic.get('km',0.0)):.3f}k)" for ic in self.ics].index(ic_label)
        except Exception:
            ic_idx = -1
        if ic_idx < 0:
            self._show_error("오류", "선택한 IC/JCT를 찾지 못했습니다.")
            return
        ramps = self.ics[ic_idx].get('ramps') or []
        # 대상 RAMP 찾기
        r_idx = -1
        try:
            r_idx = [str(r.get('name') or '') for r in ramps].index(ramp_name)
        except Exception:
            r_idx = -1
        if r_idx < 0:
            self._show_error("오류", "선택한 램프를 찾지 못했습니다.")
            return
        # 램프 길이 내부로 스냅 + 클램프
        try:
            ramp_len = float(ramps[r_idx].get('end') or 0.0)
        except Exception:
            ramp_len = 0.0
        s_local = km_floor_to_grid(s, ENTRY_GRID_KM)
        e_local = km_ceil_to_grid(e, ENTRY_GRID_KM)
        s_local = max(0.0, min(ramp_len, s_local))
        e_local = max(0.0, min(ramp_len, e_local))
        if e_local <= s_local:
            self._show_error("입력 오류", "램프 구간이 올바르지 않습니다.")
            return
        # 램프 이력: 선택 차로 반영 (전차로/1~N차로)
        ent = {"start": s_local, "end": e_local, "method": method, "ts": ts, "work_date": work_date, "lane": self.input_lane.get()}
        ramps[r_idx].setdefault('entries', []).append(ent)
        self._mark_route_data_changed(route)
        # 연도 필터 목록 갱신 후 다시 그리기
        try:
            self.update_year_filter_values()
        except Exception:
            pass
        self.draw_schematic()
        _prepare_next_input_defaults()

