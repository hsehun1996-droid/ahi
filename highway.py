#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
고속도로 포장유지보수 이력 관리 프로그램
- 다중 노선 관리 기능 (노선 추가/수정/삭제)
- 노선별 모식도에 이력 표시 (IC/JCT, 램프 포함)
- 이력 입력 시 1m 단위 스냅 및 차로 수 동적 반영
- 이력 저장/불러오기 (CSV), 모식도 내보내기 (PS, PDF) 지원
- 공법, IC/JCT, DI 지수 등 상세 정보 관리 기능
필요 라이브러리: Tkinter (기본), Pillow/reportlab (PDF 내보내기 시 선택적 필요)
"""
import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser, simpledialog, Menu
import csv
import customtkinter as ctk
import math
import sqlite3
import json
import logging
import traceback
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

# 작업 폴더가 달라도 같은 폴더의 constants/utils 등을 찾도록 프로젝트 루트를 sys.path 에 추가
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
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
from mixins import (
    UIMixin, RouteMixin, ICMixin, AnalysisMixin,
    EntryMixin, IOMixin, WindowMixin, CanvasMixin,
)


class MaintenanceApp(
    UIMixin, RouteMixin, ICMixin, AnalysisMixin,
    EntryMixin, IOMixin, WindowMixin, CanvasMixin,
    ctk.CTk,
):
    def __init__(self):
        super().__init__()
        self.logger = get_logger()
        self.title("")
        self.geometry("1200x760")

        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")  # 블루 액센트

        # 디자인 스타일 설정 적용
        self._setup_modern_style()

        # 화면 해상도에 맞게 UI 스케일 적용 (_build_ui 전에 반드시 호출)
        try:
            self._apply_display_scaling()
        except Exception:
            self._current_display_scale = 1.0
            log_exception("화면 스케일 초기화 실패")

        # 앱 아이콘 설정 (고속도로 모티프). 외부 아이콘 파일이 있으면 우선 사용
        try:
            self._set_app_icon()
        except Exception:
            log_exception("앱 아이콘 설정 실패")

        # 상태값
        self.route_name = tk.StringVar(value=DEFAULT_ROUTE_NAME)
        self.route_start_km = tk.DoubleVar(value=DEFAULT_START_KM)
        self.route_end_km = tk.DoubleVar(value=DEFAULT_END_KM)

        self.input_start_km = tk.StringVar(value="36.00")
        self.input_end_km = tk.StringVar(value="36.20")
        self.input_method = tk.StringVar(value="절삭 덧씌우기")
        # 대상(본선/IC·JCT 램프) 및 선택된 IC/RAMP
        self.input_target = tk.StringVar(value="본선")
        self.input_ic = tk.StringVar(value="")
        self.input_ramp = tk.StringVar(value="")
        self.input_direction = tk.StringVar(value=DIRECTIONS[0])
        self.input_lane = tk.StringVar(value=LANES[0])
        self.input_work_date = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))

        self.ics = []  # IC 목록: [{name, km, ramps, route}, ...]

        # 멀티 노선 관리
        self.routes = []  # 각 항목: {name, start_km, end_km, directions, entries, canvas, status_lbl, tab_frame, canvas_item_to_index}
        self.current_route_index = -1
        self.canvas_to_route_index = {}

        # 돋보기 모드
        self.magnifier_mode = False
        self.magnifier_tooltip = None
        self._magnifier_cur_path = None

        # 이력입력 모드 (본부 / 지사)
        self.entry_mode = "본부"
        self.form_title_lbl = None   # 헤더 라벨 참조
        self.form_toggle_btn = None  # 토글 버튼 참조
        self.form_card = None
        self.form_header_divider = None
        self.form_mode_badge = None
        self.form_mode_desc_lbl = None
        self.brand_logo_img = None
        self._global_data_version = 0
        self._dashboard_options_cache = None
        self._dashboard_options_cache_version = -1
        self._dashboard_summary_cache_key = None
        self._dashboard_summary_cache_values = None
        self.last_entry_by_mode = {
            "본부": {"method": self.input_method.get(), "direction": self.input_direction.get(),
                   "lane": self.input_lane.get(), "work_date": self.input_work_date.get()},
            "지사": {"method": "", "direction": self.input_direction.get(),
                   "lane": self.input_lane.get(), "work_date": self.input_work_date.get()},
        }

        # 보기/필터 상태
        self.filter_year = tk.StringVar(value="모두")  # 메인 화면의 연도 필터
        self.view_year_stack = tk.BooleanVar(value=False)
        # 모식도 보기 모드: "기본", "HPCI"
        self.view_mode = tk.StringVar(value="기본")

        # 포장상태불량률 보기 옵션
        self.view_hpci = tk.BooleanVar(value=False)
        self.view_di = tk.BooleanVar(value=False)
        self.view_aar = tk.BooleanVar(value=False)
        self.view_rd = tk.BooleanVar(value=False)
        self.view_iri = tk.BooleanVar(value=False)

        self._build_ui()
        self.draw_schematic()
        # 실행 시 자동 불러오기(동일 폴더 내 *_maintenance_history.csv)
        try:
            self.auto_load_csvs_on_start()
        except Exception:
            # 자동 불러오기는 실패해도 앱 구동에는 영향 없도록 무시
            log_exception("시작 시 자동 로드 실패")
        # 실행 시 창 크기 넓게 설정 (가로/세로 화면의 약 85%)
        try:
            self.after(50, self.set_initial_window_size)
        except Exception:
            log_exception("초기 창 크기 설정 실패")
        # 자동 불러오기 후 노선이 하나도 없으면 기본 노선을 생성
        if not self.routes:
            self.add_route(DEFAULT_ROUTE_NAME, DEFAULT_START_KM, DEFAULT_END_KM)

        # 단축키: o → 돋보기 모드 토글 (bind_all로 포커스 무관하게 전역 등록)
        try:
            self.bind_all("<KeyPress-o>", self.toggle_magnifier_mode)
            self.bind_all("<KeyPress-O>", self.toggle_magnifier_mode)
            self.bind_all("<Escape>", self.exit_magnifier_mode)
        except Exception:
            log_exception("전역 단축키 바인딩 실패")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)


# --- 전역 함수 ---
def load_custom_methods(base_dir):
    """프로그램 시작 전, all_routes_methods.csv 파일을 읽어 METHOD_STYLES를 업데이트합니다."""
    load_method_settings(base_dir)


if __name__ == "__main__":
    # 앱 실행 전에 공법을 먼저 불러옵니다.
    try:
        load_custom_methods(get_user_data_dir())
        app = MaintenanceApp()
        app.mainloop()
    except Exception:
        error_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "startup_error.log")
        with open(error_log_path, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        raise
