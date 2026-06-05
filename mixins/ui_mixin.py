# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - UI 구성, 팝업, 대시보드, 공법/하자 설정
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
from dropdown_widget import ButtonDropdown


class ModernDropdown(ctk.CTkFrame):
    """A lightweight custom dropdown with a modern popup list."""

    def __init__(self, parent, **kwargs):
        values = list(kwargs.pop("values", []))
        variable = kwargs.pop("variable", None)
        command = kwargs.pop("command", None)
        width = kwargs.pop("width", 140)
        height = kwargs.pop("height", 34)
        corner_radius = kwargs.pop("corner_radius", 12)
        border_width = kwargs.pop("border_width", 1)
        border_color = kwargs.pop("border_color", "#C8D7EA")
        fg_color = kwargs.pop("fg_color", "#FFFFFF")
        hover_color = kwargs.pop("hover_color", "#EEF4FB")
        text_color = kwargs.pop("text_color", "#213547")
        font = kwargs.pop("font", None)
        state = kwargs.pop("state", "normal")
        anchor = kwargs.pop("anchor", "w")
        dropdown_fg_color = kwargs.pop("dropdown_fg_color", "#FFFFFF")
        dropdown_hover_color = kwargs.pop("dropdown_hover_color", "#E8F1FB")
        dropdown_text_color = kwargs.pop("dropdown_text_color", text_color)
        dropdown_font = kwargs.pop("dropdown_font", font)
        self._button_color = kwargs.pop("button_color", fg_color)
        self._button_hover_color = kwargs.pop("button_hover_color", hover_color)
        self._values = [str(v) for v in values]
        self._command = command
        self._state = state
        self._anchor = anchor
        self._font = font
        self._width = width
        self._height = height
        self._corner_radius = corner_radius
        self._border_width = border_width
        self._border_color = border_color
        self._fg_color = fg_color
        self._hover_color = hover_color
        self._text_color = text_color
        self._dropdown_fg_color = dropdown_fg_color
        self._dropdown_hover_color = dropdown_hover_color
        self._dropdown_text_color = dropdown_text_color
        self._dropdown_font = dropdown_font
        self._popup = None
        self._popup_frame = None
        self._value_buttons = []
        self._is_hovered = False
        self._is_open = False
        self._hover_border_color = "#9DB4CC"
        self._active_border_color = "#7D9FC3"
        self._active_fg_color = "#F7FAFD"

        super().__init__(
            parent,
            width=width,
            height=height,
            corner_radius=corner_radius,
            fg_color=fg_color,
            border_width=border_width,
            border_color=border_color,
            **kwargs,
        )
        self.grid_propagate(False)
        self.pack_propagate(False)

        self._variable = variable if variable is not None else tk.StringVar(master=self, value="")
        self._variable_trace_id = self._variable.trace_add("write", self._sync_from_variable)

        initial_value = self._variable.get().strip()
        if initial_value:
            self._current_value = initial_value
        elif self._values:
            self._current_value = self._values[0]
            self._variable.set(self._current_value)
        else:
            self._current_value = ""

        self._label = ctk.CTkLabel(
            self,
            text=self._format_display_text(self._current_value),
            text_color=text_color,
            font=font,
            anchor=anchor,
        )
        self._label.pack(side="left", fill="both", expand=True, padx=(14, 6))

        self._chevron = ctk.CTkLabel(
            self,
            text="▼",
            width=18,
            text_color="#6F8094",
            font=font,
        )
        self._chevron.pack(side="right", padx=(4, 14))

        for widget in (self, self._label, self._chevron):
            widget.bind("<Button-1>", self._toggle_popup)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)

        self._refresh_visual_state()

    def _format_display_text(self, value):
        return value or "선택"

    def _sync_from_variable(self, *_args):
        value = self._variable.get()
        if value == self._current_value:
            return
        self._current_value = value
        self._label.configure(text=self._format_display_text(value))
        self._refresh_popup_selection()

    def _refresh_visual_state(self):
        if self._state == "disabled":
            frame_fg = "#F3F6FA"
            label_color = "#98A4B3"
            border_color = "#D9E2EC"
            chevron_color = "#B0BAC6"
            cursor = "arrow"
        else:
            frame_fg = self._active_fg_color if self._is_open else "#F8FBFE" if self._is_hovered else self._fg_color
            label_color = self._text_color
            border_color = self._active_border_color if self._is_open else self._hover_border_color if self._is_hovered else self._border_color
            chevron_color = "#496A8E" if self._is_open else "#6F8094"
            cursor = "hand2"

        ctk.CTkFrame.configure(self, fg_color=frame_fg, border_color=border_color)
        self._label.configure(text_color=label_color)
        self._chevron.configure(text_color=chevron_color, text="▴" if self._is_open else "▾")
        for widget in (self, self._label, self._chevron):
            try:
                widget.configure(cursor=cursor)
            except Exception:
                pass

    def _on_enter(self, _event=None):
        if self._state == "disabled":
            return
        self._is_hovered = True
        self._refresh_visual_state()

    def _on_leave(self, _event=None):
        if self._state == "disabled":
            return
        if self._popup and self._popup.winfo_exists():
            return
        self._is_hovered = False
        self._refresh_visual_state()

    def _toggle_popup(self, _event=None):
        if self._state == "disabled":
            return
        if self._popup and self._popup.winfo_exists():
            self._close_popup()
        else:
            self._open_popup()

    def _open_popup(self):
        if self._popup and self._popup.winfo_exists():
            return

        top = self.winfo_toplevel()
        self.update_idletasks()

        popup = ctk.CTkToplevel(top)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(fg_color="transparent")
        popup.bind("<FocusOut>", lambda _e: self._close_popup())
        popup.bind("<Escape>", lambda _e: self._close_popup())
        popup.bind("<Button-1>", self._close_on_outside_click)
        self._popup = popup
        self._is_open = True
        self._refresh_visual_state()

        option_count = max(len(self._values), 1)
        visible_rows = min(option_count, 8)
        popup_height = visible_rows * max(self._height, 34) + 18
        popup_width = max(self.winfo_width(), self._width, 160)
        pos_x = self.winfo_rootx()
        pos_y = self.winfo_rooty() + self.winfo_height() + 6
        popup.geometry(f"{popup_width}x{popup_height}+{pos_x}+{pos_y}")

        shadow = ctk.CTkFrame(
            popup,
            fg_color="#D5DFEA",
            corner_radius=max(self._corner_radius + 4, 16),
            border_width=0,
        )
        shadow.pack(fill="both", expand=True, padx=(2, 0), pady=(2, 0))
        container = ctk.CTkFrame(
            shadow,
            fg_color=self._dropdown_fg_color,
            corner_radius=max(self._corner_radius + 3, 15),
            border_width=1,
            border_color="#DCE5EF",
        )
        container.place(relx=0, rely=0, relwidth=1, relheight=1, x=-2, y=-2)
        self._popup_frame = ctk.CTkScrollableFrame(
            container,
            fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color="#D8E3EE",
            scrollbar_button_hover_color="#C4D5E7",
        )
        self._popup_frame.pack(fill="both", expand=True, padx=7, pady=7)

        self._build_popup_values()
        popup.focus_force()

    def _build_popup_values(self):
        if not self._popup_frame:
            return

        for child in self._popup_frame.winfo_children():
            child.destroy()
        self._value_buttons.clear()

        values = self._values or [""]
        for value in values:
            is_selected = value == self._current_value
            button = ctk.CTkButton(
                self._popup_frame,
                text=self._format_option_text(value, is_selected),
                height=max(self._height, 34),
                anchor="w",
                corner_radius=11,
                fg_color="#EAF1F8" if is_selected else "transparent",
                hover_color=self._dropdown_hover_color,
                text_color=self._dropdown_text_color,
                font=self._dropdown_font,
                border_width=0,
                command=lambda selected=value: self._select_value(selected),
            )
            button.pack(fill="x", pady=2, padx=1)
            self._value_buttons.append((value, button))

    def _format_option_text(self, value, is_selected):
        prefix = "✓  " if is_selected and value else "   "
        return f"{prefix}{value or '선택 안 함'}"

    def _refresh_popup_selection(self):
        for value, button in self._value_buttons:
            is_selected = value == self._current_value
            button.configure(
                text=self._format_option_text(value, is_selected),
                fg_color="#EAF1F8" if is_selected else "transparent",
            )

    def _select_value(self, value):
        self.set(value)
        self._close_popup()
        if callable(self._command):
            self._command(value)

    def _close_popup(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        self._popup_frame = None
        self._value_buttons.clear()
        self._is_open = False
        self._is_hovered = False
        self._refresh_visual_state()

    def _close_on_outside_click(self, event):
        if event.widget is self._popup:
            self._close_popup()

    def get(self):
        return self._variable.get()

    def set(self, value):
        value = "" if value is None else str(value)
        if value and self._values and value not in self._values:
            self._values.append(value)
        self._current_value = value
        self._label.configure(text=self._format_display_text(value))
        self._variable.set(value)
        self._refresh_popup_selection()

    def configure(self, require_redraw=False, **kwargs):
        if not hasattr(self, "_values"):
            return super().configure(require_redraw=require_redraw, **kwargs)

        if "values" in kwargs:
            self._values = [str(v) for v in kwargs.pop("values") or []]
            if self._values:
                current = self.get()
                if current not in self._values:
                    self.set(self._values[0])
            else:
                self.set("")
            self._build_popup_values()
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
            self._refresh_visual_state()
        if "width" in kwargs:
            self._width = kwargs["width"]
        if "height" in kwargs:
            self._height = kwargs["height"]
        if "fg_color" in kwargs:
            self._fg_color = kwargs["fg_color"]
        if "border_color" in kwargs:
            self._border_color = kwargs["border_color"]
        if "font" in kwargs:
            self._font = kwargs["font"]
            self._label.configure(font=self._font)
            self._chevron.configure(font=self._font)
        if "text_color" in kwargs:
            self._text_color = kwargs["text_color"]
        if "dropdown_fg_color" in kwargs:
            self._dropdown_fg_color = kwargs.pop("dropdown_fg_color")
        if "dropdown_hover_color" in kwargs:
            self._dropdown_hover_color = kwargs.pop("dropdown_hover_color")
        if "dropdown_text_color" in kwargs:
            self._dropdown_text_color = kwargs.pop("dropdown_text_color")
        if "dropdown_font" in kwargs:
            self._dropdown_font = kwargs.pop("dropdown_font")
        if "corner_radius" in kwargs:
            self._corner_radius = kwargs["corner_radius"]
        if "border_width" in kwargs:
            self._border_width = kwargs["border_width"]

        super().configure(require_redraw=require_redraw, **kwargs)
        self._refresh_visual_state()

    config = configure

    def cget(self, attribute_name):
        if attribute_name == "values":
            return list(self._values)
        if attribute_name == "command":
            return self._command
        if attribute_name == "state":
            return self._state
        return super().cget(attribute_name)

    def destroy(self):
        self._close_popup()
        try:
            if self._variable_trace_id:
                self._variable.trace_remove("write", self._variable_trace_id)
        except Exception:
            pass
        super().destroy()


class UIMixin:
    def _install_customtkinter_default_fonts(self):
        family = getattr(self, "font_family", "Malgun Gothic")
        default_sizes = {
            "CTkLabel": 12,
            "CTkButton": 12,
            "CTkEntry": 12,
            "CTkCheckBox": 12,
            "CTkRadioButton": 12,
            "CTkTextbox": 12,
            "CTkSegmentedButton": 11,
        }

        if getattr(ctk, "_highway_default_font_patch", None) == family:
            return

        patch_targets = [
            ctk.CTkLabel,
            ctk.CTkButton,
            ctk.CTkEntry,
            ctk.CTkCheckBox,
            ctk.CTkRadioButton,
            ctk.CTkTextbox,
            ctk.CTkSegmentedButton,
        ]

        for cls in patch_targets:
            if not hasattr(cls, "_highway_original_init"):
                cls._highway_original_init = cls.__init__

            original_init = cls._highway_original_init
            size = default_sizes.get(cls.__name__, 12)

            def _wrapped_init(widget_self, *args, __orig=original_init, __size=size, **kwargs):
                if kwargs.get("font") is None:
                    kwargs["font"] = (family, __size)
                __orig(widget_self, *args, **kwargs)

            cls.__init__ = _wrapped_init

        ctk._highway_default_font_patch = family

    def _get_brand_logo_candidates(self):
        """Return candidate logo paths for development and packaged runs."""
        candidates = [
            os.path.join(get_user_data_dir(), "concept art", "로고.JPG"),
            resource_path(os.path.join("concept art", "로고.JPG")),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "concept art", "로고.JPG"),
            os.path.join(os.getcwd(), "concept art", "로고.JPG"),
        ]
        seen = set()
        resolved = []
        for path in candidates:
            normalized = os.path.normcase(os.path.abspath(path))
            if normalized in seen:
                continue
            seen.add(normalized)
            resolved.append(path)
        return resolved

    def _find_brand_logo_path(self):
        for path in self._get_brand_logo_candidates():
            if os.path.exists(path):
                return path
        return None

    def _get_brand_icon_candidates(self):
        candidates = [
            os.path.join(get_user_data_dir(), "concept art", "logo.ico"),
            resource_path(os.path.join("concept art", "logo.ico")),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "concept art", "logo.ico"),
            os.path.join(os.getcwd(), "concept art", "logo.ico"),
            os.path.join(get_user_data_dir(), "conceptart_icon.ico"),
            resource_path("highway_icon.ico"),
        ]
        seen = set()
        resolved = []
        for path in candidates:
            normalized = os.path.normcase(os.path.abspath(path))
            if normalized in seen:
                continue
            seen.add(normalized)
            resolved.append(path)
        return resolved

    def _find_brand_icon_path(self):
        for path in self._get_brand_icon_candidates():
            if os.path.exists(path):
                return path
        return None

    def _ensure_brand_icon_ico(self):
        """Build a reusable .ico from the conceptart logo when possible."""
        existing_icon_path = self._find_brand_icon_path()
        if existing_icon_path and os.path.basename(existing_icon_path).lower() == "logo.ico":
            return existing_icon_path

        logo_path = self._find_brand_logo_path()
        if not (_PDF_LIBS_AVAILABLE and logo_path):
            return existing_icon_path

        icon_cache_path = os.path.join(get_user_data_dir(), "conceptart_icon.ico")
        try:
            logo_mtime = os.path.getmtime(logo_path)
            icon_mtime = os.path.getmtime(icon_cache_path) if os.path.exists(icon_cache_path) else -1
            if icon_mtime < logo_mtime:
                logo_img = Image.open(logo_path)
                logo_img = ImageOps.exif_transpose(logo_img).convert("RGBA")
                canvas_img = Image.new("RGBA", (256, 256), (255, 255, 255, 0))
                logo_img.thumbnail((256, 256))
                paste_x = (256 - logo_img.size[0]) // 2
                paste_y = (256 - logo_img.size[1]) // 2
                canvas_img.paste(logo_img, (paste_x, paste_y), logo_img)
                canvas_img.save(icon_cache_path, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
            return icon_cache_path
        except Exception:
            log_exception("브랜드 로고 ICO 생성 실패")
            return existing_icon_path

    def _apply_window_icon(self, window):
        return False

    def _schedule_window_icon(self, window):
        pass

    def _setup_modern_style(self):
        """애플(macOS) 스타일의 현대적인 디자인 테마 설정"""
        try:
            self._configure_app_font_family()
        except Exception:
            self.font_family = "맑은 고딕"

        try:
            self._install_customtkinter_default_fonts()
        except Exception:
            pass

        # CustomTkinter는 자체 테마를 사용하므로 ttk 스타일은 Treeview 등 일부 위젯에만 적용
        style = ttk.Style()
        style.theme_use("clam")
        
        # Treeview 라이트 모드 스타일링
        style.configure("Treeview",
                        background="#FFFFFF",
                        foreground="#213547",
                        fieldbackground="#FFFFFF",
                        borderwidth=0,
                        rowheight=28)
        style.map('Treeview', background=[('selected', '#D9EAFE')],
                              foreground=[('selected', '#123A63')])

        style.configure("Treeview.Heading",
                        background="#EAF1F9",
                        foreground=TITLE_TEXT,
                        relief="flat",
                        font=(self.font_family, 10, "bold"))
        style.map("Treeview.Heading", background=[('active', '#DCE7F5')])

        try:
            self.option_add("*Font", (self.font_family, 11))
        except Exception:
            pass

    def _create_styled_menu(self, parent=None, font_size=12):
        return tk.Menu(
            parent or self,
            tearoff=0,
            bg="#FFFFFF",
            fg=TITLE_TEXT,
            activebackground="#E8F1FB",
            activeforeground=TITLE_TEXT,
            relief="flat",
            bd=1,
            activeborderwidth=0,
            font=(self.font_family, font_size),
        )

    def _show_menu_below_widget(self, widget, menu):
        self.update_idletasks()
        menu.tk_popup(widget.winfo_rootx(), widget.winfo_rooty() + widget.winfo_height() + 4)

    def _create_dropdown_button(self, parent, text, width=120, height=30, **kwargs):
        config = {
            "text": f"{text}  ▾",
            "width": width,
            "height": height,
            "fg_color": "#FFFFFF",
            "hover_color": "#E9F1FA",
            "text_color": TITLE_TEXT,
            "border_width": 1,
            "border_color": "#C8D7EA",
            "corner_radius": 10,
            "font": (self.font_family, 12),
            "anchor": "w",
        }
        config.update(kwargs)
        return ctk.CTkButton(parent, **config)

    def _create_button(self, parent, **kwargs):
        kwargs.setdefault("font", (self.font_family, 12))
        return ctk.CTkButton(parent, **kwargs)

    def _create_styled_combobox(self, parent, **kwargs):
        config = {
            "height": 30,
            "corner_radius": 10,
            "border_width": 1,
            "border_color": "#C8D7EA",
            "fg_color": "#FFFFFF",
            "button_color": "#FFFFFF",
            "button_hover_color": "#E9F1FA",
            "text_color": TITLE_TEXT,
            "dropdown_fg_color": "#FFFFFF",
            "dropdown_hover_color": "#EEF4FB",
            "dropdown_text_color": TITLE_TEXT,
            "font": (self.font_family, 12),
            "dropdown_font": (self.font_family, 12),
        }
        config.update(kwargs)
        return ButtonDropdown(parent, **config)

    def _create_styled_option_menu(self, parent, **kwargs):
        config = {
            "height": 30,
            "corner_radius": 10,
            "border_width": 1,
            "border_color": "#C8D7EA",
            "fg_color": "#FFFFFF",
            "button_color": "#FFFFFF",
            "button_hover_color": "#E9F1FA",
            "text_color": TITLE_TEXT,
            "dropdown_fg_color": "#FFFFFF",
            "dropdown_hover_color": "#EEF4FB",
            "dropdown_text_color": TITLE_TEXT,
            "font": (self.font_family, 12),
            "dropdown_font": (self.font_family, 12),
        }
        config.update(kwargs)
        return ButtonDropdown(parent, **config)

    def _configure_app_font_family(self):
        """프로젝트 fonts 폴더의 Noto Sans KR를 우선 등록하고 기본 UI 글꼴로 사용."""
        fonts_dir = os.path.join(get_user_data_dir(), "fonts")
        if not os.path.isdir(fonts_dir):
            # 패키징 시 번들 내(_MEIPASS) 경로 폴백
            fonts_dir = resource_path("fonts")
        preferred_family = "Noto Sans KR"
        fallback_families = ["Noto Sans KR Medium", "Noto Sans KR Regular", "맑은 고딕", "Malgun Gothic", "Arial"]
        registered_paths = []

        if os.name == "nt" and os.path.isdir(fonts_dir):
            FR_PRIVATE = 0x10
            windll = getattr(ctypes, "windll", None)
            gdi32 = getattr(windll, "gdi32", None) if windll else None
            add_font = getattr(gdi32, "AddFontResourceExW", None)
            if add_font:
                for font_name in (
                    "NotoSansKR-Regular.ttf",
                    "NotoSansKR-Medium.ttf",
                    "NotoSansKR-Bold.ttf",
                    "NotoSansKR-Light.ttf",
                    "NotoSansKR-Black.ttf",
                    "NotoSansKR-Thin.ttf",
                ):
                    font_path = os.path.join(fonts_dir, font_name)
                    if os.path.exists(font_path):
                        try:
                            if add_font(font_path, FR_PRIVATE, 0):
                                registered_paths.append(font_path)
                        except Exception:
                            pass

        try:
            available_families = set(tkfont.families())
        except Exception:
            available_families = set()

        for family in [preferred_family] + fallback_families:
            if family in available_families:
                self.font_family = family
                self._registered_font_paths = registered_paths
                return

        self.font_family = "맑은 고딕"
        self._registered_font_paths = registered_paths

    # ---------- 팝업 헬퍼 메서드 (CTkMessagebox / Fallback) ----------
    def _show_info(self, title, message):
        if CTkMessagebox:
            CTkMessagebox(title=title, message=message, icon="info")
        else:
            messagebox.showinfo(title, message)

    def _create_popup_window(self, parent=None):
        dlg = ctk.CTkToplevel(parent or self)
        try:
            dlg.configure(fg_color="#FFFFFF")
        except Exception:
            pass
        try:
            self._apply_window_icon(dlg)
        except Exception:
            log_exception("팝업 아이콘 설정 실패")
        self._schedule_window_icon(dlg)
        return dlg

    def _create_close_button(self, parent, command, width=80, height=28, **kwargs):
        config = {
            "text": "닫기",
            "command": command,
            "width": width,
            "height": height,
            "fg_color": "#FFFFFF",
            "hover_color": "#E8F1FB",
            "text_color": "#000000",
            "border_width": 1,
            "border_color": "#C8D7EA",
        }
        config.update(kwargs)
        return self._create_button(parent, **config)

    def _show_priority_complete_popup(self, title, message, parent=None):
        image_path = os.path.join(get_user_data_dir(), "concept art", "길통이.JPG")
        if not os.path.exists(image_path):
            # 패키징 시 번들 내(_MEIPASS) 경로 폴백
            image_path = resource_path(os.path.join("concept art", "길통이.JPG"))
        if not _PDF_LIBS_AVAILABLE or not os.path.exists(image_path):
            self._show_info(title, message)
            return

        dlg = self._create_popup_window(parent or self)
        dlg.title("")
        dlg.geometry("420x155")
        dlg.resizable(False, False)
        dlg.transient(parent or self)
        dlg.lift()
        dlg.focus_force()
        dlg.grab_set()

        frame = ctk.CTkFrame(dlg, fg_color="#FFFFFF")
        frame.pack(fill="both", expand=True, padx=18, pady=16)

        body = ctk.CTkFrame(frame, fg_color="#FFFFFF")
        body.pack(fill="both", expand=True)

        image_holder = ctk.CTkFrame(body, width=64, height=64, corner_radius=32, fg_color="#FFFFFF")
        image_holder.pack(side="left", padx=(0, 14), pady=(4, 0))
        image_holder.pack_propagate(False)

        try:
            img = Image.open(image_path)
            img = ImageOps.exif_transpose(img)
            img.thumbnail((44, 44))
            popup_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            img_lbl = ctk.CTkLabel(image_holder, text="", image=popup_img)
            img_lbl.image = popup_img
            img_lbl.pack(expand=True)
        except Exception:
            pass

        text_frame = ctk.CTkFrame(body, fg_color="#FFFFFF")
        text_frame.pack(side="left", fill="both", expand=True)

        ctk.CTkLabel(
            text_frame,
            text=title,
            anchor="w",
            justify="left",
            font=(self.font_family, 17, "bold"),
            text_color=TITLE_TEXT
        ).pack(fill="x", pady=(2, 6))

        ctk.CTkLabel(
            text_frame,
            text=message,
            justify="left",
            anchor="w",
            wraplength=300,
            font=(self.font_family, 12)
        ).pack(fill="x")

        self._create_button(
            frame,
            text="확인",
            width=96,
            command=dlg.destroy,
            fg_color=PRIMARY_BLUE,
            hover_color=PRIMARY_BLUE_HOVER
        ).pack(pady=(14, 0))

        dlg.bind("<Return>", lambda _e: dlg.destroy())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())

    def _show_error(self, title, message):
        get_logger().error("%s: %s", title, message)
        if CTkMessagebox:
            CTkMessagebox(title=title, message=message, icon="cancel")
        else:
            messagebox.showerror(title, message)

    def _show_warning(self, title, message):
        get_logger().warning("%s: %s", title, message)
        if CTkMessagebox:
            CTkMessagebox(title=title, message=message, icon="warning")
        else:
            messagebox.showwarning(title, message)

    def _load_brand_logo(self, parent):
        """한국도로공사 로고와 제목을 상단 브랜드 영역에 표시합니다."""
        logo_path = self._find_brand_logo_path()
        if logo_path:
            try:
                img = Image.open(logo_path)
                logo_h = 30
                logo_w = max(1, int(img.size[0] * (logo_h / float(img.size[1])))) if img.size[1] else 30
                self.brand_logo_img = ctk.CTkImage(light_image=img, dark_image=img, size=(logo_w, logo_h))
                ctk.CTkLabel(parent, image=self.brand_logo_img, text="").pack(side="left", padx=(0, 8))
            except Exception:
                log_exception("공법 설정 CSV 저장 실패")

        ctk.CTkLabel(parent, text="한국도로공사",
                     font=(self.font_family, 16, "bold"),
                     text_color=TITLE_TEXT).pack(side="left", anchor="w")

    def _ask_yes_no(self, title, message):
        if CTkMessagebox:
            return CTkMessagebox(title=title, message=message, icon="question", option_1="예", option_2="아니오").get() == "예"
        else:
            return messagebox.askyesno(title, message)

    def _ask_string(self, title, prompt):
        # CTkInputDialog는 customtkinter에 내장되어 있음
        dialog = ctk.CTkInputDialog(text=prompt, title=title)
        return dialog.get_input()

    # ---------- UI 구성 ----------
    def _build_ui(self):
        # 전체 배경
        root = ctk.CTkFrame(self, fg_color=APP_BG)
        root.pack(fill="both", expand=True)

        # --- 상단 네비게이션 바 ---
        nav_bar = ctk.CTkFrame(root, fg_color=NAV_BG, height=52, corner_radius=0)
        nav_bar.pack(side="top", fill="x", padx=12, pady=(8, 0))
        # 하단 구분선
        ctk.CTkFrame(root, height=1, fg_color=NAV_BORDER, corner_radius=0).pack(side="top", fill="x", padx=14)

        nav_btn_kwargs = {
            "fg_color": "transparent",
            "text_color": NAV_TEXT,
            "hover_color": "#E9F1FA",
            "height": 36,
            "corner_radius": 12,
            "font": (self.font_family, 14),
            "border_width": 0,
        }

        def _popup_menu(btn, items):
            """버튼 아래에 팝업 메뉴 표시. items: [(label, cmd)] or None(구분선)"""
            menu = self._create_styled_menu(btn, font_size=11)
            for item in items:
                if item is None:
                    menu.add_separator()
                else:
                    lbl, cmd = item
                    menu.add_command(label=lbl, command=cmd)
            def _show(e=None):
                self._show_menu_below_widget(btn, menu)
            btn.configure(command=_show)

        brand = ctk.CTkFrame(nav_bar, fg_color="transparent")
        brand.pack(side="left", padx=(12, 14), pady=8)
        self._load_brand_logo(brand)

        # 파일 메뉴
        _file_btn = self._create_button(nav_bar, text="파일  ▾", width=99, **nav_btn_kwargs)
        _file_btn.pack(side="left", padx=(0, 3), pady=8)
        _popup_menu(_file_btn, [
            ("전체 저장", self.on_menu_save_route),
            None,
            ("전체 이력 엑셀 저장", self.on_export_all_to_excel),
        ])

        # 관리 메뉴
        _mgmt_btn = self._create_button(nav_bar, text="관리  ▾", width=99, **nav_btn_kwargs)
        _mgmt_btn.pack(side="left", padx=3, pady=8)
        _popup_menu(_mgmt_btn, [
            ("노선 관리", self.on_manage_routes),
            ("IC/JCT 관리", self.on_manage_ic),
        ])

        # 포장상태 메뉴
        _cond_btn = self._create_button(nav_bar, text="포장상태  ▾", width=127, **nav_btn_kwargs)
        _cond_btn.pack(side="left", padx=3, pady=8)
        _popup_menu(_cond_btn, [
            ("DI 지수",   self.on_manage_di),
            ("HPCI 등급", self.on_manage_hpci),
            ("AAR 등급",  self.on_manage_aar),
            ("RD 등급",   self.on_manage_rd),
            ("IRI 등급",  self.on_manage_iri),
        ])

        # 구분선
        ctk.CTkFrame(nav_bar, width=1, height=18, fg_color=NAV_BORDER).pack(side="left", padx=10, pady=14)

        # 사업계획 (강조) - 개량 우선순위 산정 + 사업계획 작성/확정
        self._create_button(nav_bar, text="사업계획",
                      command=self.on_business_plan,
                      fg_color=PRIMARY_BLUE, hover_color=PRIMARY_BLUE_HOVER,
                      text_color="#FFFFFF", width=120, height=36, corner_radius=10,
                      font=(self.font_family, 14, "bold")).pack(side="left", padx=3, pady=8)

        # 운영계획변경 (한글 양식 작성)
        self._create_button(nav_bar, text="운영계획변경",
                      command=self.on_operation_plan_change,
                      fg_color="#2F855A", hover_color="#276749",
                      text_color="#FFFFFF", width=140, height=36, corner_radius=10,
                      font=(self.font_family, 14, "bold")).pack(side="left", padx=3, pady=8)

        # 하자발생 우려구간
        self._create_button(nav_bar, text="하자발생 우려구간",
                      command=self.on_defect_risk,
                      fg_color=ACCENT_RED, hover_color=ACCENT_RED_HOVER,
                      text_color="#FFFFFF", width=172, height=36, corner_radius=10,
                      font=(self.font_family, 14, "bold")).pack(side="left", padx=3, pady=8)

        # 화면 크기 설정 버튼 (우측 끝)
        self._create_button(nav_bar, text="⚙ 화면 크기",
                      command=self.on_display_settings,
                      width=110, **nav_btn_kwargs).pack(side="right", padx=(3, 12), pady=8)

        # 메인 컨텐츠 영역
        content = ctk.CTkFrame(root, fg_color="transparent")
        content.pack(fill="both", expand=True)

        # 좌측 컨트롤 패널 (스크롤 가능 - 창 크기가 작아도 범례가 잘리지 않음)
        left_outer = ctk.CTkFrame(content, width=390, corner_radius=0, fg_color=APP_BG)
        left_outer.pack(side="left", fill="y", padx=15, pady=15)
        left_outer.pack_propagate(False)
        left = ctk.CTkScrollableFrame(left_outer, fg_color=APP_BG, scrollbar_button_color="#C8D7EA",
                                      scrollbar_button_hover_color="#9FB8D8")
        left.pack(fill="both", expand=True)

        # 우측 콘텐츠 영역
        right = ctk.CTkFrame(content, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True, padx=(0, 15), pady=15)

        # ── 스타일 상수 ──────────────────────────────────────────────────────────
        CARD_COLOR          = CARD_BG
        CARD_RADIUS         = 12
        PAD_X               = 16
        PAD_Y               = 14
        WIDGET_PAD_Y        = 5
        INPUT_BORDER_COLOR  = "#C8D7EA"
        DIVIDER_COLOR       = "#E2EAF4"
        BTN_GHOST_FG        = "#F7FAFD"
        BTN_GHOST_BORDER    = "#C9D7E7"
        BTN_GHOST_BORDER_WIDTH = 1
        BTN_GHOST_TEXT      = TITLE_TEXT
        BTN_GHOST_HOVER     = "#E8F1FB"
        BTN_MAIN_HEIGHT     = 34

        def create_header(parent, text):
            hf = ctk.CTkFrame(parent, fg_color="transparent")
            hf.pack(fill="x", padx=PAD_X, pady=(PAD_Y, 4))
            ctk.CTkLabel(hf, text=text, font=(self.font_family, 14, "bold"),
                         text_color=TITLE_TEXT, anchor="w").pack(fill="x")
            ctk.CTkFrame(hf, height=1, fg_color=DIVIDER_COLOR).pack(fill="x", pady=(4, 0))

        # 1. 이력 입력 카드
        form = ctk.CTkFrame(left, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS,
                            border_width=1, border_color="#2C5282")
        self.form_card = form
        form.pack(fill="x", pady=(0, 8))
        # 헤더: 제목 + 토글 버튼 (본부↔지사)
        hf_main = ctk.CTkFrame(form, fg_color="transparent")
        hf_main.pack(fill="x", padx=PAD_X, pady=(PAD_Y, 0))
        title_row = ctk.CTkFrame(hf_main, fg_color="transparent")
        title_row.pack(fill="x")
        title_left = ctk.CTkFrame(title_row, fg_color="transparent")
        title_left.pack(side="left", fill="x", expand=True)
        self.form_title_lbl = ctk.CTkLabel(
            title_left, text="이력입력(본부)",
            font=(self.font_family, 16, "bold"), anchor="w")
        self.form_title_lbl.pack(side="left")
        self.form_mode_badge = ctk.CTkLabel(
            title_left, text="본부 모드",
            width=74, height=24, corner_radius=12,
            font=(self.font_family, 11, "bold"),
            text_color="#EBF8FF", fg_color="#2C5282")
        self.form_mode_badge.pack(side="left", padx=(10, 0))
        self.form_toggle_btn = self._create_button(
            title_row, text="지사", width=52, height=26,
            command=self.toggle_entry_mode,
            fg_color="#D7E6F7", hover_color="#C7DBF3", text_color=TITLE_TEXT, corner_radius=8)
        self.form_toggle_btn.pack(side="right")
        self.form_mode_desc_lbl = ctk.CTkLabel(
            hf_main,
            text="본부 확정 이력을 등록합니다. 전체 이력, 우선순위, 하자분석에 직접 반영됩니다.",
            anchor="w", justify="left",
            font=(self.font_family, 11), text_color="#3B6E9E")
        self.form_mode_desc_lbl.pack(fill="x", pady=(6, 0))
        self.form_header_divider = ctk.CTkFrame(hf_main, height=1, fg_color="#2C5282")
        self.form_header_divider.pack(fill="x", pady=(5, 0))

        form_inner = ctk.CTkFrame(form, fg_color="transparent")
        form_inner.pack(fill="x", padx=PAD_X, pady=(0, PAD_Y))

        ctk.CTkLabel(form_inner, text="시작(km)").grid(row=0, column=0, sticky="w", pady=WIDGET_PAD_Y)
        ctk.CTkEntry(form_inner, textvariable=self.input_start_km, width=70, height=28, border_color=INPUT_BORDER_COLOR).grid(row=0, column=1, sticky="w", padx=(5,0), pady=WIDGET_PAD_Y)
        ctk.CTkLabel(form_inner, text="끝(km)").grid(row=0, column=2, sticky="w", padx=(10,0), pady=WIDGET_PAD_Y)
        ctk.CTkEntry(form_inner, textvariable=self.input_end_km, width=70, height=28, border_color=INPUT_BORDER_COLOR).grid(row=0, column=3, sticky="w", padx=(5,0), pady=WIDGET_PAD_Y)

        ctk.CTkLabel(form_inner, text="공법").grid(row=1, column=0, sticky="w", pady=WIDGET_PAD_Y)
        self.method_cb = self._create_styled_combobox(form_inner, variable=self.input_method, values=list(METHOD_STYLES.keys()), width=100, height=28, state="readonly", border_color=INPUT_BORDER_COLOR)
        self.method_cb.grid(row=1, column=1, sticky="w", padx=(5,0), pady=WIDGET_PAD_Y)

        ctk.CTkLabel(form_inner, text="시공날짜").grid(row=1, column=2, sticky="w", padx=(10,0), pady=WIDGET_PAD_Y)
        ctk.CTkEntry(form_inner, textvariable=self.input_work_date, width=100, height=28, border_color=INPUT_BORDER_COLOR).grid(row=1, column=3, sticky="w", padx=(5,0), pady=WIDGET_PAD_Y)

        ctk.CTkLabel(form_inner, text="방향").grid(row=2, column=0, sticky="w", pady=WIDGET_PAD_Y)
        self.dir_cb = self._create_styled_combobox(form_inner, variable=self.input_direction, values=DIRECTIONS, width=84, height=28, state="readonly", border_color=INPUT_BORDER_COLOR)
        self.dir_cb.grid(row=2, column=1, sticky="w", padx=(5,0), pady=WIDGET_PAD_Y)

        ctk.CTkLabel(form_inner, text="차로").grid(row=2, column=2, sticky="w", padx=(10,0), pady=WIDGET_PAD_Y)
        self.lane_cb = self._create_styled_combobox(form_inner, variable=self.input_lane, values=LANES, width=70, height=28, state="readonly", border_color=INPUT_BORDER_COLOR)
        self.lane_cb.grid(row=2, column=3, sticky="w", padx=(5,0), pady=WIDGET_PAD_Y)

        btn_row = ctk.CTkFrame(form_inner, fg_color="transparent")
        btn_row.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(10,0))
        
        self._create_button(btn_row, text="이력 추가", command=self.on_add_entry, height=32, corner_radius=8,
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self._create_button(btn_row, text="세부 이력 확인", command=self.on_open_detail_table, height=32, corner_radius=8,
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(side="right", fill="x", expand=True, padx=(5, 0))

        # 2. 데이터 관리 (버튼 그룹)
        file_card = ctk.CTkFrame(left, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS)
        file_card.pack(fill="x", pady=(0, 8))
        create_header(file_card, "데이터 관리")
        
        file_inner = ctk.CTkFrame(file_card, fg_color="transparent")
        file_inner.pack(fill="x", padx=PAD_X, pady=(0, PAD_Y))
        
        self._create_button(file_inner, text="구조물 관리", command=self.on_manage_structures, height=32, corner_radius=8, 
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(fill="x", pady=3)
        
        row_csv = ctk.CTkFrame(file_inner, fg_color="transparent")
        row_csv.pack(fill="x", pady=3)
        self._create_button(row_csv, text="CSV 저장", command=self.on_save_csv, height=32, corner_radius=8,
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(side="left", fill="x", expand=True, padx=(0,3))
        self._create_button(row_csv, text="CSV 불러오기", command=self.on_load_csv, height=32, corner_radius=8,
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(side="right", fill="x", expand=True, padx=(3,0))

        row_excel = ctk.CTkFrame(file_inner, fg_color="transparent")
        row_excel.pack(fill="x", pady=3)
        self._create_button(row_excel, text="Excel 불러오기", command=self.on_load_excel, height=32, corner_radius=8,
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(fill="x")
        
        self._create_button(file_inner, text="모식도 내보내기(PDF)", command=self.on_export_pdf, height=32, corner_radius=8, 
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(fill="x", pady=(3,0))

        # 4. 보기/필터
        view = ctk.CTkFrame(left, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS)
        view.pack(fill="x", pady=(0, 8))
        create_header(view, "보기/필터")

        view_inner = ctk.CTkFrame(view, fg_color="transparent")
        view_inner.pack(fill="x", padx=PAD_X, pady=(0, PAD_Y))

        ctk.CTkLabel(view_inner, text="연도").grid(row=0, column=0, sticky="w", pady=WIDGET_PAD_Y)
        self.year_cb = self._create_styled_combobox(view_inner, variable=self.filter_year, values=["모두"], state="readonly", width=60, height=28, border_color=INPUT_BORDER_COLOR)
        self.year_cb.grid(row=0, column=1, sticky="w", padx=(5,0), pady=WIDGET_PAD_Y)
        def _apply_year(*_):
            self.on_apply_year_filter()
        self.year_cb.configure(command=_apply_year)
        self._create_button(view_inner, text="적용", command=self.on_apply_year_filter, width=60, height=28, corner_radius=8,
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).grid(row=0, column=2, sticky="w", padx=(5,0), pady=WIDGET_PAD_Y)

        ctk.CTkLabel(view_inner, text="모식도 보기").grid(row=1, column=0, sticky="w", pady=WIDGET_PAD_Y)
        mode_frame = ctk.CTkFrame(view_inner, fg_color="transparent")
        mode_frame.grid(row=1, column=1, columnspan=3, sticky="w", pady=WIDGET_PAD_Y, padx=(10, 0))

        def on_view_mode_change():
            # '이력' 모드일 때는 포장상태 체크박스 해제
            if self.view_mode.get() == "기본":
                self.view_aar.set(False)
                self.view_hpci.set(False)
                self.view_di.set(False)
                self.view_rd.set(False)
                self.view_iri.set(False)
            self.draw_schematic()

        def on_condition_view_change():
            # 포장상태 항목이 선택되면 보기 모드를 'CONDITION'으로 강제 변경
            if self.view_hpci.get() or self.view_di.get() or self.view_aar.get() or self.view_rd.get() or self.view_iri.get():
                self.view_mode.set("CONDITION")
            self.draw_schematic()

        ctk.CTkRadioButton(mode_frame, text="이력", variable=self.view_mode, value="기본", command=on_view_mode_change, radiobutton_width=16, radiobutton_height=16).pack(side="left", padx=(0, 10))
        
        # 포장상태불량률 드롭다운 메뉴
        condition_menubutton = self._create_dropdown_button(mode_frame, "포장상태불량률", width=116, height=30)
        condition_menu = self._create_styled_menu(condition_menubutton)
        condition_menu.add_checkbutton(label="HPCI", variable=self.view_hpci, command=on_condition_view_change)
        condition_menu.add_checkbutton(label="DI", variable=self.view_di, command=on_condition_view_change)
        condition_menu.add_checkbutton(label="AAR등급", variable=self.view_aar, command=on_condition_view_change)
        condition_menu.add_checkbutton(label="RD", variable=self.view_rd, command=on_condition_view_change)
        condition_menu.add_checkbutton(label="IRI", variable=self.view_iri, command=on_condition_view_change)
        condition_menubutton.configure(command=lambda: self._show_menu_below_widget(condition_menubutton, condition_menu))
        condition_menubutton.pack(side="left")

        # 메뉴 항목 클릭 시 메뉴가 닫히지 않도록 이벤트 전파를 중단합니다.
        def stay_open(event):
            return "break"

        condition_menu.bind("<ButtonRelease-1>", stay_open)

        # 메뉴 외부를 클릭했을 때 메뉴를 닫는 로직
        def on_click_outside(event):
            try:
                widget = event.widget
                if isinstance(widget, str):
                    try:
                        widget = self.nametowidget(widget)
                    except Exception:
                        return
                
                master = getattr(widget, 'master', None)
                if master not in (condition_menu, condition_menubutton) and widget not in (condition_menu, condition_menubutton):
                    condition_menu.unpost()
            except Exception:
                pass
        self.bind_all("<Button-1>", on_click_outside, add="+")
        condition_menu.config(tearoff=0) # tearoff 기능을 비활성화하여 점선이 보이지 않게 합니다.

        # 범례: 좌측 파일 카드 아래로 이동 및 공법 추가 기능
        self.legend_frame = ctk.CTkFrame(left, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS)
        self.legend_frame.pack(fill="x", pady=(0, 8))
        create_header(self.legend_frame, "범례")

        controls = ctk.CTkFrame(self.legend_frame, fg_color="transparent")
        controls.pack(fill="x", padx=PAD_X, pady=(0, 5))
        self._create_button(controls, text="공법 관리", command=self.on_manage_methods, height=24, corner_radius=6,
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(side="left")
        self._create_button(controls, text="하자기간 설정", command=self.on_warranty_settings, height=24, corner_radius=6,
                      fg_color=BTN_GHOST_FG, hover_color=BTN_GHOST_HOVER, text_color=BTN_GHOST_TEXT,
                      border_width=BTN_GHOST_BORDER_WIDTH, border_color=BTN_GHOST_BORDER).pack(side="left", padx=(6, 0))

        legend_container = ctk.CTkFrame(self.legend_frame, fg_color="transparent")
        legend_container.pack(fill="x", padx=PAD_X, pady=5)

        self.legend_list = ctk.CTkScrollableFrame(legend_container, fg_color="transparent", height=160)
        self.legend_list.pack(fill="x")

        self.legend_hint_lbl = ctk.CTkLabel(
            self.legend_frame, text="얇은 가로선은 차로 구분, 세로 점선은 100m",
            text_color=MUTED_TEXT, font=(self.font_family, 10))
        self.legend_hint_lbl.pack(fill="x", padx=PAD_X, pady=(0, 10))

        # 대시보드 요약 바
        self._build_dashboard(right)

        # 우측 탭뷰
        self.notebook = ctk.CTkTabview(
            right,
            corner_radius=CARD_RADIUS,
            fg_color="#F3F8FD",
            segmented_button_fg_color="#D9E6F3",
            segmented_button_selected_color=PRIMARY_BLUE,
            segmented_button_unselected_color="#6F8FB3",
            segmented_button_selected_hover_color=PRIMARY_BLUE_HOVER,
            segmented_button_unselected_hover_color="#5D81AA",
            text_color="#FFFFFF"
        )
        self.notebook.pack(fill="both", expand=True)
        self.notebook.configure(command=self.on_tab_changed)


        # 초기 범례 렌더링
        self.render_legend()

        # 초기 보기 모드 설정 적용
        on_view_mode_change()
        self._apply_entry_mode_ui()

        # 기본 노선은 자동 불러오기 후, 노선이 하나도 없을 때에만 생성 (초기 화면 깨짐 방지용)

    def _set_app_icon(self):
        if os.name == "nt":
            import ctypes
            def _remove_icon():
                try:
                    hwnd = self.winfo_id()
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, 0)  # WM_SETICON ICON_SMALL
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 1, 0)  # WM_SETICON ICON_BIG
                except Exception:
                    pass
            self.after(100, _remove_icon)

    def _build_dashboard(self, parent):
        """우측 패널 상단에 요약 통계 카드 바를 생성합니다."""
        dash = ctk.CTkFrame(parent, fg_color="#EDF4FB", corner_radius=14, border_width=1, border_color=CARD_BORDER)
        dash.pack(fill="x", pady=(0, 6))
        self._dash_frame = dash

        # 필터 바
        filter_bar = ctk.CTkFrame(dash, fg_color="transparent")
        filter_bar.pack(fill="x", padx=8, pady=(6, 2))

        ctk.CTkLabel(filter_bar, text="범위", font=(self.font_family, 11),
                     text_color=MUTED_TEXT).pack(side="left", padx=(0, 4))
        self._dash_scope = tk.StringVar(value="전체 노선")
        ctk.CTkSegmentedButton(
            filter_bar, values=["전체 노선", "현재 노선"],
            variable=self._dash_scope, command=lambda _: self.refresh_dashboard(),
            width=180, height=24, font=(self.font_family, 11),
            selected_color=PRIMARY_BLUE, selected_hover_color=PRIMARY_BLUE_HOVER,
            unselected_color="#D8E5F2", unselected_hover_color="#C6D8EC",
            text_color="white", text_color_disabled=TITLE_TEXT
        ).pack(side="left")

        ctk.CTkLabel(filter_bar, text="연도", font=(self.font_family, 11),
                     text_color=MUTED_TEXT).pack(side="left", padx=(12, 4))
        self._dash_year = tk.StringVar(value="모두")
        self._dash_year_cb = self._create_styled_combobox(
            filter_bar, variable=self._dash_year, values=["모두"],
            width=60, height=24, state="readonly",
            command=lambda _: self.refresh_dashboard()
        )
        self._dash_year_cb.pack(side="left")

        ctk.CTkLabel(filter_bar, text="공법", font=(self.font_family, 11),
                     text_color=MUTED_TEXT).pack(side="left", padx=(10, 4))
        self._dash_method = tk.StringVar(value="모두")
        self._dash_method_cb = self._create_styled_combobox(
            filter_bar, variable=self._dash_method, values=["모두"],
            width=60, height=24, state="readonly",
            command=lambda _: self.refresh_dashboard()
        )
        self._dash_method_cb.pack(side="left")

        # 통계 카드 5개
        inner = ctk.CTkFrame(dash, fg_color="transparent")
        inner.pack(fill="x", padx=8, pady=(0, 4))
        for i in range(5):
            inner.columnconfigure(i, weight=1)

        card_configs = [
            ("노선 수",      "routes"),
            ("총 연장",      "total_km"),
            ("시공 연장",    "year_work"),
            ("이력 건수",    "total_entries"),
            ("불량포장 연장", "bad_pavement"),
        ]

        self._dash_labels = {}
        for col, (title, key) in enumerate(card_configs):
            card = ctk.CTkFrame(inner, fg_color="#2A2A2B", corner_radius=8)
            card.configure(fg_color="#FFFFFF", border_width=1, border_color=CARD_BORDER)
            card.grid(row=0, column=col, padx=4, sticky="ew")
            ctk.CTkLabel(card, text=title, font=(self.font_family, 11),
                         text_color=MUTED_TEXT).pack(pady=(4, 0))
            color = ACCENT_RED if key == "bad_pavement" else PRIMARY_BLUE
            val_lbl = ctk.CTkLabel(card, text="—", font=(self.font_family, 13, "bold"),
                                   text_color=color)
            val_lbl.pack(pady=(0, 4))
            self._dash_labels[key] = val_lbl

    def refresh_dashboard(self):
        """대시보드 통계 카드의 값을 현재 필터에 맞게 갱신합니다."""
        if not hasattr(self, "_dash_labels"):
            return
        try:
            scope_val  = self._dash_scope.get()  if hasattr(self, "_dash_scope")  else "전체 노선"
            sel_year   = self._dash_year.get()   if hasattr(self, "_dash_year")   else "모두"
            sel_method = self._dash_method.get() if hasattr(self, "_dash_method") else "모두"

            # 대상 노선 결정
            if scope_val == "현재 노선" and 0 <= self.current_route_index < len(self.routes):
                target_routes = [self.routes[self.current_route_index]]
            else:
                target_routes = self.routes

            # 드롭다운 선택지 갱신 (전체 노선 기준) - 데이터 버전 기준 캐시
            if self._dashboard_options_cache_version != self._global_data_version or self._dashboard_options_cache is None:
                all_years: set[str] = set()
                all_methods: set[str] = set()
                for r in self.routes:
                    for e in r.get("entries", []):
                        wd = str(e.get("work_date", ""))
                        if len(wd) >= 4:
                            all_years.add(wd[:4])
                        m = e.get("method", "")
                        if m:
                            all_methods.add(m)
                self._dashboard_options_cache = {
                    "years": ["모두"] + sorted(all_years, reverse=True),
                    "methods": ["모두"] + sorted(all_methods),
                }
                self._dashboard_options_cache_version = self._global_data_version

            year_values = self._dashboard_options_cache["years"]
            method_values = self._dashboard_options_cache["methods"]

            if hasattr(self, "_dash_year_cb"):
                self._dash_year_cb.configure(values=year_values)
                if sel_year not in year_values:
                    self._dash_year.set("모두"); sel_year = "모두"
            if hasattr(self, "_dash_method_cb"):
                self._dash_method_cb.configure(values=method_values)
                if sel_method not in method_values:
                    self._dash_method.set("모두"); sel_method = "모두"

            summary_key = (
                self._global_data_version,
                scope_val,
                self.current_route_index if scope_val == "현재 노선" else -1,
                sel_year,
                sel_method,
            )
            if self._dashboard_summary_cache_key == summary_key and self._dashboard_summary_cache_values is not None:
                n_routes, total_km, year_work_m, total_entries, bad_pavement_m = self._dashboard_summary_cache_values
            else:
                n_routes = len(target_routes)
                total_km = sum(
                    abs(r.get("end_km", 0) - r.get("start_km", 0))
                    for r in target_routes
                )
                year_work_m = 0.0
                total_entries = 0
                for r in target_routes:
                    for e in r.get("entries", []):
                        wd = str(e.get("work_date", ""))
                        method = e.get("method", "")
                        if sel_year != "모두" and not wd.startswith(sel_year):
                            continue
                        if sel_method != "모두" and method != sel_method:
                            continue
                        total_entries += 1
                        year_work_m += abs(e.get("end", 0) - e.get("start", 0)) * 1000

                bad_segments: set[tuple] = set()
                for r in target_routes:
                    hpci_data = r.get("hpci_data", {})
                    for (d, l, skm_str), year_map in hpci_data.items():
                        if not isinstance(year_map, dict) or not year_map:
                            continue
                        if sel_year != "모두":
                            grade = year_map.get(sel_year)
                        else:
                            grade = year_map.get(sorted(year_map.keys())[-1])
                        try:
                            if grade is not None and int(grade) >= 5:
                                bad_segments.add((id(r), d, l, skm_str))
                        except (ValueError, TypeError):
                            pass
                bad_pavement_m = len(bad_segments) * GRID_KM * 1000
                self._dashboard_summary_cache_key = summary_key
                self._dashboard_summary_cache_values = (
                    n_routes, total_km, year_work_m, total_entries, bad_pavement_m
                )

            self._dash_labels["routes"].configure(text=str(n_routes))
            self._dash_labels["total_km"].configure(text=f"{total_km:.1f} km")
            self._dash_labels["year_work"].configure(text=f"{year_work_m:,.0f} m")
            self._dash_labels["total_entries"].configure(text=f"{total_entries:,} 건")
            self._dash_labels["bad_pavement"].configure(text=f"{bad_pavement_m:,.0f} m")
        except Exception:
            pass

    def render_legend(self):
        for child in self.legend_list.winfo_children():
            child.destroy()

        # 카테고리별로 그룹화
        from collections import defaultdict
        groups: dict[str, list] = defaultdict(list)
        uncategorized = []
        for name, style in METHOD_STYLES.items():
            cat = METHOD_CATEGORY_MAP.get(name, "")
            if cat:
                groups[cat].append((name, style))
            else:
                uncategorized.append((name, style))

        def _add_method_row(name, style, indent=0):
            row = ctk.CTkFrame(self.legend_list, fg_color="transparent")
            row.pack(anchor="w", pady=1, padx=(indent, 0))
            ctk.CTkLabel(row, text="■", text_color=style.get("fill", "#718096"),
                         font=(self.font_family, 13)).pack(side="left")
            ctk.CTkLabel(row, text=f" {name}",
                         font=(self.font_family, 12)).pack(side="left")

        # 카테고리 헤더 + 소속 공법
        for cat in sorted(groups.keys()):
            hdr = ctk.CTkFrame(self.legend_list, fg_color="transparent")
            hdr.pack(anchor="w", pady=(5, 1))
            ctk.CTkLabel(hdr, text=f"▸ {cat}",
                         font=(self.font_family, 12, "bold"),
                         text_color="#A0AEC0").pack(side="left")
            for name, style in groups[cat]:
                _add_method_row(name, style, indent=12)

        # 카테고리 없는 공법
        if uncategorized and groups:
            hdr = ctk.CTkFrame(self.legend_list, fg_color="transparent")
            hdr.pack(anchor="w", pady=(5, 1))
            ctk.CTkLabel(hdr, text="▸ 기타",
                         font=(self.font_family, 12, "bold"),
                         text_color="#A0AEC0").pack(side="left")
        for name, style in uncategorized:
            _add_method_row(name, style, indent=12 if groups else 0)

    def update_method_combobox_values(self):
        self.method_cb.configure(values=list(METHOD_STYLES.keys()))

    def on_warranty_settings(self):
        """카테고리별 하자기간·보증률 설정 다이얼로그."""
        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry("480x420")
        dlg.resizable(False, True)

        header = ctk.CTkFrame(dlg, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(header, text="카테고리별 하자기간·보증률 설정",
                     font=(self.font_family, 13, "bold")).pack(side="left")
        ctk.CTkLabel(header, text="(미설정 카테고리는 기본값 3년 적용)",
                     font=(self.font_family, 10), text_color="gray60").pack(side="left", padx=(8, 0))

        # 컬럼 헤더
        col_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        col_frame.pack(fill="x", padx=12)
        ctk.CTkLabel(col_frame, text="카테고리", width=160, anchor="w",
                     font=(self.font_family, 11, "bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(col_frame, text="하자기간(년)", width=100, anchor="center",
                     font=(self.font_family, 11, "bold")).grid(row=0, column=1)
        ctk.CTkLabel(col_frame, text="보증률(%)", width=100, anchor="center",
                     font=(self.font_family, 11, "bold")).grid(row=0, column=2)

        scroll = ctk.CTkScrollableFrame(dlg, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=12, pady=4)

        cats = sorted(set(METHOD_CATEGORY_MAP.values()))
        row_vars = {}  # {cat: (period_var, rate_var)}

        for i, cat in enumerate(cats):
            cur = CATEGORY_WARRANTY.get(cat, {})
            period_var = tk.StringVar(value=str(cur.get("period", 3)))
            rate_var   = tk.StringVar(value=str(cur.get("rate", 100.0)))
            row_vars[cat] = (period_var, rate_var)

            ctk.CTkLabel(scroll, text=cat, width=160, anchor="w",
                         font=(self.font_family, 11)).grid(row=i, column=0, sticky="w", pady=3)
            ctk.CTkEntry(scroll, textvariable=period_var, width=90,
                         justify="center").grid(row=i, column=1, padx=8, pady=3)
            ctk.CTkEntry(scroll, textvariable=rate_var, width=90,
                         justify="center").grid(row=i, column=2, padx=8, pady=3)

        if not cats:
            ctk.CTkLabel(scroll, text="등록된 카테고리가 없습니다.\n공법 관리에서 카테고리를 먼저 설정해 주세요.",
                         text_color="gray60").pack(pady=20)

        def on_save():
            for cat, (pv, rv) in row_vars.items():
                try:
                    period = int(pv.get())
                    if period < 1: period = 1
                except ValueError:
                    period = 3
                try:
                    rate = float(rv.get())
                    if not (0 <= rate <= 100): rate = 100.0
                except ValueError:
                    rate = 100.0
                CATEGORY_WARRANTY[cat] = {"period": period, "rate": rate}
            # 파일 저장
            try:
                base_dir = get_user_data_dir()
                fpath = os.path.join(base_dir, "warranty_settings.csv")
                with open(fpath, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f)
                    w.writerow(["category", "period", "rate"])
                    for cat, info in CATEGORY_WARRANTY.items():
                        w.writerow([cat, info.get("period", 3), info.get("rate", 100.0)])
            except Exception:
                pass
            dlg.destroy()

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(fill="x", padx=12, pady=(4, 12))
        self._create_button(btn_row, text="저장", command=on_save, width=100).pack(side="right")
        self._create_close_button(btn_row, dlg.destroy, width=100).pack(side="right", padx=(0, 6))

    def on_manage_methods(self):
        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry("520x340")

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=10)

        cols = ("name", "category", "color")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=8)
        tree.heading("name",     text="공법명")
        tree.heading("category", text="카테고리")
        tree.heading("color",    text="색상")
        tree.column("name",     width=180)
        tree.column("category", width=130)
        tree.column("color",    width=80, anchor="center")
        tree.pack(fill="both", expand=True)

        def refresh_tree():
            for item in tree.get_children():
                tree.delete(item)
            for name, style in METHOD_STYLES.items():
                color = style.get('fill', '#000000')
                cat   = METHOD_CATEGORY_MAP.get(name, "")
                iid = tree.insert("", "end", values=(name, cat, color))
                tree.tag_configure(iid, background=color)

        refresh_tree()

        def _open_method_edit_dialog(old_name=None):
            edit_dlg = self._create_popup_window(dlg)
            is_edit = old_name is not None
            edit_dlg.title("")
            edit_dlg.transient(dlg)
            edit_dlg.grab_set()

            edit_body = ctk.CTkFrame(edit_dlg, fg_color="transparent")
            edit_body.pack(fill="both", expand=True, padx=10, pady=10)

            var_name = tk.StringVar(value=old_name if is_edit else "")
            initial_color = METHOD_STYLES[old_name]['fill'] if is_edit else "#4A5568"
            var_color = tk.StringVar(value=initial_color)
            initial_cat = METHOD_CATEGORY_MAP.get(old_name, "") if is_edit else ""
            var_cat = tk.StringVar(value=initial_cat)

            ctk.CTkLabel(edit_body, text="공법명").grid(row=0, column=0, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(edit_body, textvariable=var_name, width=160).grid(row=0, column=1, sticky="w", padx=4, pady=4)

            ctk.CTkLabel(edit_body, text="카테고리").grid(row=1, column=0, sticky="w", padx=4, pady=4)
            existing_cats = sorted(set(METHOD_CATEGORY_MAP.values()))
            cat_cb = self._create_styled_combobox(edit_body, variable=var_cat,
                                                  values=["(없음)"] + existing_cats,
                                                  width=160)
            cat_cb.grid(row=1, column=1, sticky="w", padx=4, pady=4)

            ctk.CTkLabel(edit_body, text="색상").grid(row=2, column=0, sticky="w", padx=4, pady=4)
            color_row = ctk.CTkFrame(edit_body, fg_color="transparent")
            color_row.grid(row=2, column=1, sticky="w", padx=4, pady=4)
            swatch = ctk.CTkLabel(color_row, text="■", text_color=var_color.get(), font=(self.font_family, 20))
            swatch.pack(side="left", padx=(0, 6))

            def pick_color():
                _, hexcolor = colorchooser.askcolor(initialcolor=var_color.get(), parent=edit_dlg, title="색상 선택")
                if hexcolor:
                    var_color.set(hexcolor)
                    swatch.configure(text_color=hexcolor)

            self._create_button(color_row, text="색상 선택", command=pick_color, width=80).pack(side="left")

            def on_save():
                new_name = var_name.get().strip()
                new_color = var_color.get().strip()

                if not new_name:
                    self._show_error("입력 오류", "공법명을 입력해 주세요.")
                    return

                if new_name != old_name and new_name in METHOD_STYLES:
                    self._show_error("중복", "이미 존재하는 공법명입니다.")
                    return

                new_cat = var_cat.get().strip()
                if new_cat == "(없음)":
                    new_cat = ""

                if is_edit:
                    if old_name != new_name:
                        for route in self.routes:
                            for entry in route.get("entries", []):
                                if entry.get("method") == old_name:
                                    entry["method"] = new_name
                        for ic in getattr(self, 'ics', []):
                            for ramp in ic.get('ramps', []):
                                for rentry in ramp.get('entries', []):
                                    if rentry.get("method") == old_name:
                                        rentry["method"] = new_name
                        METHOD_STYLES[new_name] = METHOD_STYLES.pop(old_name)
                        # 카테고리 키도 이전
                        old_cat = METHOD_CATEGORY_MAP.pop(old_name, "")
                        if old_cat:
                            METHOD_CATEGORY_MAP[new_name] = old_cat

                    METHOD_STYLES[new_name]['fill'] = new_color
                else:
                    METHOD_STYLES[new_name] = {'fill': new_color}

                # 카테고리 반영
                if new_cat:
                    METHOD_CATEGORY_MAP[new_name] = new_cat
                else:
                    METHOD_CATEGORY_MAP.pop(new_name, None)

                self.update_method_combobox_values()
                self.render_legend()
                self.draw_schematic()
                refresh_tree()
                edit_dlg.destroy()

            btn_frame = ctk.CTkFrame(edit_dlg, fg_color="transparent")
            btn_frame.pack(fill="x", padx=10, pady=10)
            self._create_button(btn_frame, text="저장", command=on_save, width=80).pack(side="right")
            self._create_button(btn_frame, text="취소", command=edit_dlg.destroy, width=80, fg_color="transparent", border_width=1).pack(side="right", padx=5)

        def on_add():
            _open_method_edit_dialog()

        def on_edit():
            selected = tree.selection()
            if not selected:
                self._show_info("선택 필요", "수정할 공법을 목록에서 선택해 주세요.")
                return
            old_name = tree.item(selected[0], "values")[0]
            _open_method_edit_dialog(old_name)

        def on_delete():
            selected = tree.selection()
            if not selected:
                self._show_info("선택 필요", "삭제할 공법을 목록에서 선택해 주세요.")
                return
            
            name_to_delete = tree.item(selected[0], "values")[0]

            is_in_use = False
            for route in self.routes:
                for entry in route.get("entries", []):
                    if entry.get("method") == name_to_delete:
                        is_in_use = True
                        break
                if is_in_use:
                    break
            
            if not is_in_use:
                for ic in getattr(self, 'ics', []):
                    for ramp in ic.get('ramps', []):
                        for rentry in ramp.get('entries', []):
                            if rentry.get("method") == name_to_delete:
                                is_in_use = True
                                break
                        if is_in_use:
                            break
                    if is_in_use:
                        break

            if is_in_use:
                self._show_error("삭제 불가", f"'{name_to_delete}' 공법은 현재 사용 중이므로 삭제할 수 없습니다.")
                return

            if self._ask_yes_no("삭제 확인", f"'{name_to_delete}' 공법을 삭제하시겠습니까?"):
                if name_to_delete in METHOD_STYLES:
                    del METHOD_STYLES[name_to_delete]
                METHOD_CATEGORY_MAP.pop(name_to_delete, None)
                
                self.update_method_combobox_values()
                self.render_legend()
                self.draw_schematic()
                refresh_tree()

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)

        self._create_button(btn_frame, text="추가", command=on_add, width=80).pack(side="left")
        self._create_button(btn_frame, text="수정", command=on_edit, width=80).pack(side="left", padx=5)
        self._create_button(btn_frame, text="삭제", command=on_delete, width=80, fg_color="#C53030", hover_color="#9B2C2C").pack(side="left")
        self._create_close_button(btn_frame, dlg.destroy, width=80).pack(side="right")

    # ---------- 멀티 노선 ----------
