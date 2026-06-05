# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - 노선 관리 (추가/수정/삭제, 탭, 필터)
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


class RouteMixin:
    def add_route(self, route_name: str, start_km: float, end_km: float, directions=None):
        # CTkTabview에 탭 추가
        self.notebook.add(route_name)
        tab = self.notebook.tab(route_name)

        # 상단 상태바
        topbar = ctk.CTkFrame(tab, fg_color="transparent")
        topbar.pack(fill="x", padx=8, pady=(8, 0))
        status_lbl = ctk.CTkLabel(topbar, text="", text_color=TITLE_TEXT)
        status_lbl.pack(side="left")

        # 캔버스
        canvas_frame = ctk.CTkFrame(tab, fg_color="#FFFFFF", corner_radius=16,
                                    border_width=1, border_color=CARD_BORDER)
        canvas_frame.pack(fill="both", expand=True, pady=(6,0), padx=8)
        canvas = tk.Canvas(canvas_frame, bg="white", height=CANVAS_WIDGET_H)
        hbar = ctk.CTkScrollbar(canvas_frame, orientation="horizontal", command=canvas.xview)
        canvas.configure(xscrollcommand=hbar.set)
        canvas.pack(fill="both", expand=True)
        hbar.pack(fill="x")

        canvas.bind("<Double-Button-1>", self.on_canvas_double_click)
        canvas.bind("<MouseWheel>", self.on_mouse_wheel_horizontal)
        canvas.bind("<Button-4>", self.on_mouse_wheel_horizontal)
        canvas.bind("<Button-5>", self.on_mouse_wheel_horizontal)
        canvas.bind("<Button-1>", lambda e: e.widget.focus_set())

        # 캔버스 배경색 다크 모드 적용
        canvas.configure(bg=BACKGROUND_FILL, highlightthickness=0)

        # 돋보기 상세 모식도 컨테이너 (초기에는 숨김 - pack 안 함)
        detail_container = ctk.CTkFrame(tab, fg_color="transparent")
        detail_header = ctk.CTkFrame(detail_container, fg_color="transparent")
        detail_header.pack(fill="x", padx=4, pady=(4, 0))
        detail_title_lbl = ctk.CTkLabel(detail_header, text="", text_color=TITLE_TEXT,
                                        font=(self.font_family, 10, "bold"))
        detail_title_lbl.pack(side="left")
        detail_canvas_frame = ctk.CTkFrame(detail_container, fg_color="#FFFFFF", corner_radius=14,
                                           border_width=1, border_color=CARD_BORDER)
        detail_canvas_frame.pack(fill="both", expand=True)
        detail_canvas = tk.Canvas(detail_canvas_frame, bg=BACKGROUND_FILL,
                                  height=DETAIL_CANVAS_H, highlightthickness=0)
        detail_hbar = ctk.CTkScrollbar(detail_canvas_frame, orientation="horizontal",
                                       command=detail_canvas.xview)
        detail_canvas.configure(xscrollcommand=detail_hbar.set)
        detail_canvas.pack(fill="both", expand=True)
        detail_hbar.pack(fill="x")
        detail_canvas.bind("<MouseWheel>", self.on_mouse_wheel_horizontal)
        detail_canvas.bind("<Button-4>", self.on_mouse_wheel_horizontal)
        detail_canvas.bind("<Button-5>", self.on_mouse_wheel_horizontal)
        detail_canvas.bind("<Double-Button-1>", self.on_detail_canvas_double_click)

        route = {
            "name": route_name,
            "start_km": float(start_km),
            "end_km": float(end_km),
            "directions": directions if directions is not None else self.generate_directions_from_name(route_name),
            "entries": [],
            "canvas": canvas,
            "hbar": hbar,
            "status_lbl": status_lbl,
            "tab_frame": tab,
            "canvas_item_to_index": {},
            "lane_count": 4,
            "ic_item_to_index": {},
            "di_target_item_to_info": {},
            "di_data": {},
            "aar_data": {},
            "hpci_data": {},
            "rd_data": {},
            "iri_data": {},
            "pavement_data": {},
            "detail_container": detail_container,
            "detail_canvas": detail_canvas,
            "detail_hbar": detail_hbar,
            "detail_title_lbl": detail_title_lbl,
            "detail_section_km": None,
            "branch_entries": [],
            "branch_item_to_index": {},
            "list_tree": None,
            "tree_item_to_index": {},
            "structures": [],
            "_last_render_signature": None,
            "_data_version": 0,
            "_year_values_cache": None,
            "_year_values_version": -1,
        }
        self.routes.append(route)
        self.canvas_to_route_index[canvas] = len(self.routes) - 1

        # 상세보기 영역을 처음부터 표시하고 안내 문구 그리기
        detail_container.pack(fill="x", pady=(4, 0))
        def _draw_detail_hint(event, _route=route, _dc=detail_canvas):
            if _route.get("detail_section_km") is not None:
                return
            _dc.delete("hint_text")
            _dc.create_text(
                event.width // 2, event.height // 2,
                text="10m 단위 상세보기를 원하시면 키보드 O를 누른 후 원하는 구간을 클릭하세요 (지사 보수현황 확인)",
                fill="#AAAAAA", font=(self.font_family, 11), anchor="center", tags="hint_text"
            )
        detail_canvas.bind("<Configure>", _draw_detail_hint, add="+")

        # 현재 탭으로 전환
        self.notebook.set(route_name)
        self.current_route_index = len(self.routes) - 1
        self.refresh_route_controls_from_route(route)
        self.draw_schematic()

    def on_manage_routes(self):
        """노선 관리(추가/수정/삭제) 대화상자를 엽니다."""
        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry("600x400")

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=10)

        # 노선 목록
        cols = ("name", "start_km", "end_km", "directions", "lanes")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=10)
        tree.heading("name", text="노선명")
        tree.heading("start_km", text="시작 이정(km)")
        tree.heading("end_km", text="끝 이정(km)")
        tree.heading("directions", text="방향")
        tree.heading("lanes", text="차로 수")
        tree.column("name", width=120)
        tree.column("start_km", width=100, anchor="center")
        tree.column("end_km", width=100, anchor="center")
        tree.column("directions", width=200)
        tree.column("lanes", width=60, anchor="center")
        tree.pack(fill="both", expand=True)

        route_map = {} # tree item id -> route index

        def refresh_tree():
            for item in tree.get_children():
                tree.delete(item)
            route_map.clear()
            for i, route in enumerate(self.routes):
                dirs = ", ".join(route.get("directions", []))
                lanes = route.get("lane_count", 4)
                values = (route["name"], f"{route['start_km']:.2f}", f"{route['end_km']:.2f}", dirs, lanes)
                item_id = tree.insert("", "end", values=values)
                route_map[item_id] = i

        refresh_tree()

        def on_add():
            # 기존 노선 추가 로직 재사용
            self.on_add_route(parent=dlg)
            # 추가 후 목록 새로고침
            refresh_tree()

        def on_edit():
            selected = tree.selection()
            if not selected:
                self._show_info("선택 필요", "수정할 노선을 목록에서 선택해 주세요.")
                return
            
            route_idx = route_map.get(selected[0])
            if route_idx is None or not (0 <= route_idx < len(self.routes)):
                return
            
            route = self.routes[route_idx]
            
            # 수정 다이얼로그 (on_add_route과 유사)
            edit_dlg = self._create_popup_window(dlg)
            edit_dlg.title("")
            edit_dlg.transient(dlg)
            edit_dlg.grab_set()

            body_edit = ctk.CTkFrame(edit_dlg, fg_color="transparent")
            body_edit.pack(fill="both", expand=True, padx=10, pady=10)

            var_name = tk.StringVar(value=route['name'])
            var_start = tk.StringVar(value=f"{route['start_km']:.2f}")
            var_end = tk.StringVar(value=f"{route['end_km']:.2f}")
            dirs = route.get('directions', ["", ""])
            var_dir1 = tk.StringVar(value=dirs[0].replace("방향", ""))
            var_dir2 = tk.StringVar(value=dirs[1].replace("방향", "") if len(dirs) > 1 else "")
            var_lanes = tk.StringVar(value=str(route.get("lane_count", 4)))

            ctk.CTkLabel(body_edit, text="노선명").grid(row=0, column=0, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(body_edit, textvariable=var_name, width=160).grid(row=0, column=1, columnspan=2, sticky="w", padx=4, pady=4)
            ctk.CTkLabel(body_edit, text="시작 이정(km)").grid(row=1, column=0, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(body_edit, textvariable=var_start, width=80).grid(row=1, column=1, sticky="w", padx=4, pady=4)
            ctk.CTkLabel(body_edit, text="끝 이정(km)").grid(row=1, column=2, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(body_edit, textvariable=var_end, width=80).grid(row=1, column=3, sticky="w", padx=4, pady=4)
            ctk.CTkLabel(body_edit, text="방향1").grid(row=2, column=0, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(body_edit, textvariable=var_dir1, width=80).grid(row=2, column=1, sticky="w", padx=4, pady=4)
            ctk.CTkLabel(body_edit, text="방향2").grid(row=2, column=2, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(body_edit, textvariable=var_dir2, width=80).grid(row=2, column=3, sticky="w", padx=4, pady=4)
            ctk.CTkLabel(body_edit, text="차로 수").grid(row=3, column=0, sticky="w", padx=4, pady=4)
            cb_lanes = self._create_styled_combobox(body_edit, variable=var_lanes, values=[str(i) for i in range(1, 7)], width=80, state="readonly")
            cb_lanes.grid(row=3, column=1, sticky="w", padx=4, pady=4)

            def do_save():
                # 값 검증 및 반영
                name = var_name.get().strip()
                try:
                    s = float(var_start.get().strip())
                    e = float(var_end.get().strip())
                except ValueError:
                    self._show_error("입력 오류", "시작/끝 이정을 숫자로 입력해 주세요.")
                    return
                if not name or e <= s:
                    self._show_error("입력 오류", "노선명과 이정을 올바르게 입력해 주세요.")
                    return
                try:
                    lanes = int(var_lanes.get())
                except ValueError:
                    self._show_error("입력 오류", "차로 수를 선택해 주세요.")
                    return

                route['name'] = name
                route['start_km'] = s
                route['end_km'] = e
                dir1 = var_dir1.get().strip()
                dir2 = var_dir2.get().strip()
                if dir1 and dir2:
                    route['directions'] = [f"{dir1}방향", f"{dir2}방향"]
                route['lane_count'] = lanes
                
                self.notebook.tab(route['tab_frame'], text=name)
                self.refresh_route_controls_from_route(route)
                self.draw_schematic()
                refresh_tree()
                edit_dlg.destroy()

            btn_frame = ctk.CTkFrame(edit_dlg, fg_color="transparent")
            btn_frame.pack(fill="x", padx=10, pady=10)
            self._create_button(btn_frame, text="저장", command=do_save, width=80).pack(side="right")
            self._create_button(btn_frame, text="취소", command=edit_dlg.destroy, width=80, fg_color="transparent", border_width=1).pack(side="right", padx=5)

        def on_delete():
            selected = tree.selection()
            if not selected:
                self._show_info("선택 필요", "삭제할 노선을 목록에서 선택해 주세요.")
                return
            
            route_idx = route_map.get(selected[0])
            if route_idx is None or not (0 <= route_idx < len(self.routes)):
                return

            route_to_delete = self.routes[route_idx]
            if self._ask_yes_no("삭제 확인", f"'{route_to_delete['name']}' 노선을 삭제하시겠습니까?\n이 노선과 관련된 모든 이력 정보가 사라집니다."):
                # 탭 제거
                self.notebook.delete(route_to_delete['name'])
                # self.routes에서 제거
                del self.routes[route_idx]
                # canvas_to_route_index 맵 재구성
                self.canvas_to_route_index = {r['canvas']: i for i, r in enumerate(self.routes)}
                
                # 현재 선택된 탭 인덱스 조정
                if self.current_route_index == route_idx:
                    if not self.routes: # 모든 탭이 삭제된 경우
                        self.current_route_index = -1
                        # 빈 화면 처리 (예: 기본 노선 추가 또는 안내 메시지)
                        self.add_route(DEFAULT_ROUTE_NAME, DEFAULT_START_KM, DEFAULT_END_KM)
                    else:
                        # 다른 탭으로 포커스 이동
                        new_idx = max(0, route_idx - 1)
                        self.notebook.set(self.routes[new_idx]['name'])
                elif self.current_route_index > route_idx:
                    self.current_route_index -= 1

                refresh_tree()
                # 메인 UI 갱신
                if self.routes:
                    current_route = self.routes[self.current_route_index]
                    self.refresh_route_controls_from_route(current_route)
                    self.draw_schematic()
                else:
                    # 모든 노선이 삭제되었을 때의 처리
                    self.route_name.set("")
                    self.route_start_km.set(0.0)
                    self.route_end_km.set(0.0)

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)
        self._create_button(btn_frame, text="추가", command=on_add, width=80).pack(side="left")
        self._create_button(btn_frame, text="수정", command=on_edit, width=80).pack(side="left", padx=5)
        self._create_button(btn_frame, text="삭제", command=on_delete, width=80, fg_color="#C53030", hover_color="#9B2C2C").pack(side="left")
        self._create_close_button(btn_frame, dlg.destroy, width=80).pack(side="right")

    def on_add_route(self, parent=None):
        owner = parent if parent is not None else self
        dlg = self._create_popup_window(owner)
        dlg.title("")
        dlg.transient(owner)
        dlg.grab_set()

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=10)

        var_name = tk.StringVar(value="")
        var_start = tk.StringVar(value=f"{DEFAULT_START_KM:.2f}")
        var_end = tk.StringVar(value=f"{DEFAULT_END_KM:.2f}")
        var_dir1 = tk.StringVar(value="")
        var_dir2 = tk.StringVar(value="")
        var_lanes = tk.StringVar(value="4")

        ctk.CTkLabel(body, text="노선명").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ent_name = ctk.CTkEntry(body, textvariable=var_name, width=160)
        ent_name.grid(row=0, column=1, columnspan=2, sticky="w", padx=4, pady=4)

        ctk.CTkLabel(body, text="시작 이정(km)").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ent_start = ctk.CTkEntry(body, textvariable=var_start, width=80)
        ent_start.grid(row=1, column=1, sticky="w", padx=4, pady=4)

        ctk.CTkLabel(body, text="끝 이정(km)").grid(row=1, column=2, sticky="w", padx=4, pady=4)
        ent_end = ctk.CTkEntry(body, textvariable=var_end, width=80)
        ent_end.grid(row=1, column=3, sticky="w", padx=4, pady=4)

        ctk.CTkLabel(body, text="방향1").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        ent_dir1 = ctk.CTkEntry(body, textvariable=var_dir1, width=80)
        ent_dir1.grid(row=2, column=1, sticky="w", padx=4, pady=4)

        ctk.CTkLabel(body, text="방향2").grid(row=2, column=2, sticky="w", padx=4, pady=4)
        ent_dir2 = ctk.CTkEntry(body, textvariable=var_dir2, width=80)
        ent_dir2.grid(row=2, column=3, sticky="w", padx=4, pady=4)

        ctk.CTkLabel(body, text="차로 수").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        cb_lanes = self._create_styled_combobox(body, variable=var_lanes, values=[str(i) for i in range(1, 7)], width=80, state="readonly")
        cb_lanes.grid(row=3, column=1, sticky="w", padx=4, pady=4)

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(fill="x", padx=10, pady=10)

        def do_add():
            name = var_name.get().strip()
            try:
                s = float(var_start.get().strip())
                e = float(var_end.get().strip())
            except Exception:
                self._show_error("입력 오류", "시작/끝 이정을 숫자로 입력해 주세요. 예: 36.00, 126.91")
                return
            if not name:
                self._show_error("입력 오류", "노선명을 입력해 주세요.")
                return
            if any(r["name"] == name for r in self.routes):
                self._show_error("중복 오류", f"'{name}' 노선이 이미 존재합니다. 다른 이름을 사용해 주세요.")
                return
            if e <= s:
                self._show_error("입력 오류", "끝 이정은 시작 이정보다 커야 합니다.")
                return
            dir1 = var_dir1.get().strip() 
            dir2 = var_dir2.get().strip()
            custom_dirs = None
            if dir1 and dir2:
                custom_dirs = [dir1 if dir1.endswith("방향") else f"{dir1}방향",
                               dir2 if dir2.endswith("방향") else f"{dir2}방향"]
            try:
                lanes = int(var_lanes.get())
            except ValueError:
                lanes = 4
            
            self.add_route(name, s, e)
            new_route = self.routes[-1]
            new_route['lane_count'] = lanes
            if custom_dirs:
                new_route['directions'] = custom_dirs
            dlg.destroy()

        self._create_button(btns, text="추가", command=do_add, width=80).pack(side="left")
        self._create_button(btns, text="취소", command=dlg.destroy, width=80, fg_color="transparent", border_width=1).pack(side="right")

    def on_tab_changed(self):
        # CTkTabview는 event 인자 없이 호출됨
        selected_name = self.notebook.get()
        # 이름으로 인덱스 찾기
        idx = -1
        for i, r in enumerate(self.routes):
            if r["name"] == selected_name:
                idx = i
                break
        if idx == -1: return

        self.current_route_index = idx
        route = self.routes[idx]
        self.refresh_route_controls_from_route(route)
        if self._needs_route_redraw(route):
            self.draw_schematic()
        else:
            self.refresh_dashboard()

    def _route_render_signature(self, route: dict):
        """탭 전환 시 불필요한 재렌더링을 피하기 위한 경량 시그니처."""
        return (
            route.get("name"),
            round(float(route.get("start_km", 0.0)), 3),
            round(float(route.get("end_km", 0.0)), 3),
            tuple(route.get("directions", [])),
            int(route.get("lane_count", 4)),
            int(route.get("_data_version", 0)),
            self.filter_year.get(),
            bool(self.view_hpci.get()),
            bool(self.view_di.get()),
            bool(self.view_aar.get()),
            bool(self.view_rd.get()),
            bool(self.view_iri.get()),
            str(self.view_mode.get()),
            bool(getattr(self, "schematic_select_mode", False)),
        )

    def _needs_route_redraw(self, route: dict) -> bool:
        canvas = route.get("canvas")
        if canvas is None:
            return True
        if not canvas.find_all():
            return True
        current_sig = self._route_render_signature(route)
        return current_sig != route.get("_last_render_signature")

    def _mark_route_data_changed(self, route: dict | None = None):
        """데이터 변경 버전을 올려 캐시를 무효화합니다."""
        self._global_data_version += 1
        self._dashboard_options_cache = None
        self._dashboard_options_cache_version = -1
        self._dashboard_summary_cache_key = None
        self._dashboard_summary_cache_values = None
        if route is not None:
            route["_data_version"] = int(route.get("_data_version", 0)) + 1
            route["_year_values_cache"] = None
            route["_year_values_version"] = -1

    def _get_route_year_values(self, route: dict):
        cached = route.get("_year_values_cache")
        cached_ver = route.get("_year_values_version", -1)
        route_ver = route.get("_data_version", 0)
        if cached is not None and cached_ver == route_ver:
            return cached

        years = set()
        try:
            for it in route.get("entries", []):
                wd = str(it.get("work_date") or "")
                if len(wd) >= 4 and wd[:4].isdigit():
                    years.add(wd[:4])

            current_route_name = str(route.get("name", ""))
            for ic in getattr(self, 'ics', []):
                if str(ic.get('route') or '') != current_route_name:
                    continue
                for r in ic.get('ramps') or []:
                    for ent in r.get('entries', []) or []:
                        wd = str(ent.get('work_date') or "")
                        if len(wd) >= 4 and wd[:4].isdigit():
                            years.add(wd[:4])

            for data_key in ['di_data', 'hpci_data', 'aar_data', 'rd_data', 'iri_data']:
                data_map = route.get(data_key, {})
                for val_map in data_map.values():
                    if isinstance(val_map, dict):
                        years.update(val_map.keys())
        except Exception:
            pass

        values = ["모두"] + sorted(years)
        route["_year_values_cache"] = values
        route["_year_values_version"] = route_ver
        return values

    def refresh_route_controls_from_route(self, route):
        # 상단 노선 정보 바인딩 업데이트
        self.route_name.set(route["name"])
        self.route_start_km.set(route["start_km"])
        self.route_end_km.set(route["end_km"])
        # 방향 콤보박스 업데이트
        self.dir_cb.configure(values=route["directions"])
        # 차로 콤보박스 업데이트
        lane_count = route.get("lane_count", 4)
        self.input_lane.set("전차로")
        self.lane_cb.configure(values=["전차로"] + [f"{i}차로" for i in range(1, lane_count + 1)])
        self.input_direction.set(route["directions"][0])
        # 범례 방향 힌트 업데이트
        try:
            dirs = route.get("directions", [])
            up = dirs[0] if dirs else "상행"
            down = dirs[1] if len(dirs) > 1 else "하행"
            self.legend_hint_lbl.configure(
                text=f"위: {up}, 아래: {down}; 얇은 가로선은 차로 구분, 세로 점선은 100m"
            )
        except Exception:
            pass
        # 연도 필터 콤보 갱신
        self.update_year_filter_values()

    def update_year_filter_values(self):
        values = ["모두"]
        if 0 <= self.current_route_index < len(self.routes):
            values = self._get_route_year_values(self.routes[self.current_route_index])
        try:
            self.year_cb.configure(values=values)
        except Exception:
            pass
        # 기존 선택이 유효하지 않으면 모두로 설정
        if self.filter_year.get() not in values:
            self.filter_year.set("모두")

    def on_apply_year_filter(self):
        # 단순히 다시 그리기
        self.draw_schematic()

    def infer_route_directions_from_entries(self, route: dict):
        """이력에 기록된 실제 방향 값들로 노선의 방향 라벨을 재설정.
        - 둘 다 존재하면 그대로 사용(순서는 처음 등장 순서)
        - 하나만 존재하면 다른 하나는 그 반대방향이 흔한 경우가 아니면 기본값으로 보완
        - 전혀 없으면 기본값 유지
        """
        try:
            seen = []
            for it in route.get("entries", []):
                d = (it.get("direction") or "").strip()
                if d and d not in seen:
                    seen.append(d)
            if len(seen) >= 2:
                route["directions"] = [seen[0], seen[1]]
            elif len(seen) == 1:
                d = seen[0]
                # 간단 규칙: '평택방향'이면 '제천방향' 보완, '제천방향'이면 '평택방향' 보완, 그 외에는 기본 두 값
                if d == DIRECTIONS[0]:
                    route["directions"] = [DIRECTIONS[0], DIRECTIONS[1]]
                elif d == DIRECTIONS[1]:
                    route["directions"] = [DIRECTIONS[1], DIRECTIONS[0]]
                else:
                    route["directions"] = [d, DIRECTIONS[1] if d == DIRECTIONS[0] else DIRECTIONS[0]]
            else:
                # 기록된 방향이 없으면 기존/기본 유지
                if not route.get("directions"):
                    route["directions"] = DIRECTIONS[:]
        except Exception:
            # 실패시 기본 유지
            if not route.get("directions"):
                route["directions"] = DIRECTIONS[:]

