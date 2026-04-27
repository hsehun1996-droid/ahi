# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - IC/구조물/상태 데이터 관리
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


class ICMixin:
    def on_manage_ic(self):
        # IC/JCT 관리 대화상자
        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.transient(self)
        dlg.grab_set()

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=10)

        # 리스트 영역
        cols = ("route", "name", "km", "ramps")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=8)
        tree.heading("route", text="노선명")
        tree.heading("name", text="IC/JCT 이름")
        tree.heading("km", text="이정(km)")
        tree.heading("ramps", text="램프(이정)")
        tree.column("route", width=180, anchor="center")
        tree.column("name", width=160, anchor="center")
        tree.column("km", width=100, anchor="center")
        tree.column("ramps", width=520, anchor="w")
        tree.pack(fill="both", expand=True)

        def format_ramps_for_list(ic_rec):
            ramps = ic_rec.get('ramps') or []
            parts = []
            for r in ramps:
                if isinstance(r, dict):
                    nm = r.get('name') or '-'
                    try:
                        km_r = float(r.get('km', 0.0))
                        parts.append(f"{nm} ({km_r:.3f}k)")
                    except Exception:
                        parts.append(str(nm))
                else:
                    parts.append(str(r))
            return ", ".join(parts) if parts else "-"

        def refresh_tree():
            for it in tree.get_children():
                tree.delete(it)
            for ic in getattr(self, 'ics', []):
                ramp_txt = format_ramps_for_list(ic)
                tree.insert("", "end", values=(ic.get('route', ''), ic.get('name', ''), f"{float(ic.get('km', 0.0)):.3f}", ramp_txt))

        refresh_tree()

        # 편집 영역
        frm = ctk.CTkFrame(dlg, fg_color="transparent")
        frm.pack(fill="x", padx=10, pady=(0,10))
        ctk.CTkLabel(frm, text="노선").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        var_route = tk.StringVar()
        route_names = [r.get("name", f"노선{idx}") for idx, r in enumerate(self.routes)] or [self.route_name.get()]
        cb_route = self._create_styled_combobox(frm, variable=var_route, values=route_names, state="readonly", width=140)
        cb_route.grid(row=0, column=1, sticky="w", padx=4, pady=4)
        if self.current_route_index >= 0 and self.current_route_index < len(route_names):
            var_route.set(route_names[self.current_route_index])

        ctk.CTkLabel(frm, text="IC/JCT 입력").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        var_name = tk.StringVar()
        ctk.CTkEntry(frm, textvariable=var_name, width=140).grid(row=1, column=1, sticky="w", padx=4, pady=4)
        ctk.CTkLabel(frm, text="이정(km)").grid(row=1, column=2, sticky="w", padx=4, pady=4)
        var_km = tk.StringVar()
        ctk.CTkEntry(frm, textvariable=var_km, width=80).grid(row=1, column=3, sticky="w", padx=4, pady=4)

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(fill="x", padx=10, pady=(0,10))

        def on_new_ic():
            try:
                name = var_name.get().strip() or "IC"
                km = float(var_km.get().strip())
            except Exception:
                self._show_error("입력 오류", "이정(km)을 숫자로 입력해 주세요. 예: 124.000")
                return
            ramps = []
            self.ics.append({"name": name, "km": km, "ramps": ramps, "route": var_route.get()})
            var_name.set("")
            var_km.set("")
            refresh_tree()
            self.draw_schematic()

        def on_delete_selected():
            try:
                sel = tree.selection()
                if not sel:
                    return
                vals = tree.item(sel[0], "values")
                route_name = vals[0]
                name = vals[1]
                km = float(vals[2])
                self.ics = [ic for ic in self.ics if not (str(ic.get('route') or '')==route_name and ic.get('name')==name and float(ic.get('km',0.0))==km)]
                refresh_tree()
                self.draw_schematic()
            except Exception:
                pass

        self._create_button(btns, text="IC/JCT 추가", command=on_new_ic, width=100).pack(side="left", padx=(0,5))
        self._create_button(btns, text="선택 삭제", command=on_delete_selected, width=100, fg_color="#C53030", hover_color="#9B2C2C").pack(side="left")

        # 더블클릭 시 수정 팝업
        def on_tree_double_click_ic(_evt=None):
            try:
                sel = tree.selection()
                if not sel:
                    return
                vals = tree.item(sel[0], "values")
                route_name = vals[0]
                name = vals[1]
                km = float(vals[2])
                target_idx = None
                for i, ic in enumerate(getattr(self, 'ics', [])):
                    if str(ic.get('route') or '')==route_name and ic.get('name')==name and abs(float(ic.get('km',0.0)) - km) < 1e-6:
                        target_idx = i
                        break
                if target_idx is None:
                    return
                ic = self.ics[target_idx]

                ed = self._create_popup_window(dlg)
                ed.title("")
                ed.transient(dlg)
                ed.grab_set()

                ebody = ctk.CTkFrame(ed, fg_color="transparent")
                ebody.pack(fill="both", expand=True, padx=10, pady=10)

                # 상단: 노선/이름/이정
                ctk.CTkLabel(ebody, text="노선").grid(row=0, column=0, sticky="w", padx=4, pady=4)
                var_eroute = tk.StringVar(value=str(ic.get('route') or route_name))
                cb_eroute = self._create_styled_combobox(ebody, variable=var_eroute, values=[r.get("name","") for r in self.routes], state="readonly", width=140)
                cb_eroute.grid(row=0, column=1, sticky="w", padx=4, pady=4)

                ctk.CTkLabel(ebody, text="IC/JCT 이름").grid(row=0, column=2, sticky="w", padx=4, pady=4)
                var_ename = tk.StringVar(value=str(ic.get('name') or ""))
                ctk.CTkEntry(ebody, textvariable=var_ename, width=140).grid(row=0, column=3, sticky="w", padx=4, pady=4)

                ctk.CTkLabel(ebody, text="이정(km)").grid(row=1, column=0, sticky="w", padx=4, pady=4)
                var_ekm = tk.StringVar(value=f"{float(ic.get('km',0.0)):.3f}")
                ctk.CTkEntry(ebody, textvariable=var_ekm, width=80).grid(row=1, column=1, sticky="w", padx=4, pady=4)

                # 램프 관리 영역
                ctk.CTkLabel(ebody, text="램프 목록", font=(self.font_family, 11, "bold")).grid(
                    row=2, column=0, columnspan=4, sticky="w", padx=4, pady=(10, 2))

                ramp_cols = ("ramp_name", "ramp_km")
                ramp_tree = ttk.Treeview(ebody, columns=ramp_cols, show="headings", height=4)
                ramp_tree.heading("ramp_name", text="램프명")
                ramp_tree.heading("ramp_km", text="이정(km)")
                ramp_tree.column("ramp_name", width=140, anchor="center")
                ramp_tree.column("ramp_km", width=90, anchor="center")
                ramp_tree.grid(row=3, column=0, columnspan=4, sticky="nsew", padx=4, pady=2)

                # 현재 램프 목록 채우기
                edit_ramps = [dict(r) if isinstance(r, dict) else {"name": str(r), "km": 0.0}
                              for r in (ic.get('ramps') or [])]

                def refresh_ramp_tree():
                    for it in ramp_tree.get_children():
                        ramp_tree.delete(it)
                    for r in edit_ramps:
                        ramp_tree.insert("", "end", values=(r.get('name',''), f"{float(r.get('km',0.0)):.3f}"))

                refresh_ramp_tree()

                # 램프 추가 입력
                ramp_add_frame = ctk.CTkFrame(ebody, fg_color="transparent")
                ramp_add_frame.grid(row=4, column=0, columnspan=4, sticky="w", padx=4, pady=2)
                var_rname = tk.StringVar()
                var_rkm = tk.StringVar()
                ctk.CTkLabel(ramp_add_frame, text="램프명").pack(side="left", padx=(0,4))
                ctk.CTkEntry(ramp_add_frame, textvariable=var_rname, width=110).pack(side="left", padx=(0,8))
                ctk.CTkLabel(ramp_add_frame, text="이정(km)").pack(side="left", padx=(0,4))
                ctk.CTkEntry(ramp_add_frame, textvariable=var_rkm, width=70).pack(side="left", padx=(0,8))

                def on_add_ramp():
                    nm = var_rname.get().strip()
                    if not nm:
                        return
                    try:
                        rkm = float(var_rkm.get().strip())
                    except Exception:
                        self._show_error("입력 오류", "이정(km)을 숫자로 입력해 주세요.")
                        return
                    edit_ramps.append({"name": nm, "km": rkm})
                    var_rname.set("")
                    var_rkm.set("")
                    refresh_ramp_tree()

                def on_del_ramp():
                    sel = ramp_tree.selection()
                    if not sel:
                        return
                    idx_r = ramp_tree.index(sel[0])
                    if 0 <= idx_r < len(edit_ramps):
                        edit_ramps.pop(idx_r)
                    refresh_ramp_tree()

                self._create_button(ramp_add_frame, text="추가", command=on_add_ramp, width=60).pack(side="left", padx=(0,4))
                self._create_button(ramp_add_frame, text="선택삭제", command=on_del_ramp, width=70,
                              fg_color="#C53030", hover_color="#9B2C2C").pack(side="left")

                ebtns = ctk.CTkFrame(ed, fg_color="transparent")
                ebtns.pack(fill="x", padx=10, pady=10)

                def on_save_e():
                    # 검증 및 저장
                    try:
                        new_km = float(var_ekm.get().strip())
                    except Exception:
                        self._show_error("입력 오류", "이정(km)을 숫자로 입력해 주세요. 예: 124.000")
                        return
                    new_name = var_ename.get().strip()
                    if not new_name:
                        self._show_error("입력 오류", "IC/JCT 이름을 입력해 주세요.")
                        return
                    new_route = var_eroute.get().strip()

                    # 저장 반영
                    self.ics[target_idx] = {"name": new_name, "km": new_km, "ramps": list(edit_ramps), "route": new_route}
                    refresh_tree()
                    self.draw_schematic()
                    ed.destroy()

                self._create_button(ebtns, text="저장", command=on_save_e, width=80).pack(side="right")
                self._create_button(ebtns, text="취소", command=ed.destroy, width=80, fg_color="transparent", border_width=1).pack(side="right", padx=(0,5))
            except Exception:
                pass

        tree.bind("<Double-1>", on_tree_double_click_ic)

    def on_manage_structures(self):
        """구조물(교량, 터널) 관리 대화상자를 엽니다."""
        if not (0 <= self.current_route_index < len(self.routes)):
            self._show_error("노선 없음", "구조물을 관리할 노선을 선택해 주세요.")
            return

        route = self.routes[self.current_route_index]

        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry("600x400")

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=10)

        cols = ("name", "type", "direction", "start_km", "length_m")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=10)
        tree.heading("name", text="구조물명")
        tree.heading("type", text="종류")
        tree.heading("direction", text="방향")
        tree.heading("start_km", text="시작 이정(km)")
        tree.heading("length_m", text="교장(m)")
        tree.column("name", width=150)
        tree.column("type", width=80, anchor="center")
        tree.column("direction", width=80, anchor="center")
        tree.column("start_km", width=100, anchor="center")
        tree.column("length_m", width=80, anchor="center")
        tree.pack(fill="both", expand=True)

        item_to_idx = {}

        def refresh_tree():
            for item in tree.get_children():
                tree.delete(item)
            item_to_idx.clear()
            for i, struct in enumerate(route.get("structures", [])):
                direction = struct.get('direction', '양방향')
                length = struct.get('length', 0.0)
                if length == 0.0 and struct.get('end') and struct.get('start'):
                     length = (struct['end'] - struct['start']) * 1000.0
                
                values = (struct['name'], struct['type'], direction, f"{struct['start']:.3f}", f"{length:.1f}")
                iid = tree.insert("", "end", values=values)
                item_to_idx[iid] = i

        refresh_tree()

        def open_edit_dialog(struct_idx=None):
            is_edit = struct_idx is not None
            struct_data = route["structures"][struct_idx] if is_edit else {}

            edit_dlg = self._create_popup_window(dlg)
            edit_dlg.title("")
            edit_dlg.transient(dlg)
            edit_dlg.grab_set()

            edit_body = ctk.CTkFrame(edit_dlg, fg_color="transparent")
            edit_body.pack(fill="both", expand=True, padx=10, pady=10)

            cur_dir = struct_data.get('direction', '양방향')
            cur_len = struct_data.get('length', 0.0)
            if cur_len == 0.0 and struct_data.get('end') and struct_data.get('start'):
                cur_len = (struct_data['end'] - struct_data['start']) * 1000.0

            var_name = tk.StringVar(value=struct_data.get('name', ''))
            var_type = tk.StringVar(value=struct_data.get('type', '교량'))
            var_dir = tk.StringVar(value=cur_dir)
            var_start = tk.StringVar(value=f"{struct_data.get('start', 0.0):.3f}")
            var_length = tk.StringVar(value=f"{cur_len:.1f}")

            ctk.CTkLabel(edit_body, text="구조물명").grid(row=0, column=0, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(edit_body, textvariable=var_name, width=160).grid(row=0, column=1, sticky="w", padx=4, pady=4)
            ctk.CTkLabel(edit_body, text="종류").grid(row=1, column=0, sticky="w", padx=4, pady=4)
            self._create_styled_combobox(edit_body, variable=var_type, values=["교량", "터널"], state="readonly", width=80).grid(row=1, column=1, sticky="w", padx=4, pady=4)
            
            ctk.CTkLabel(edit_body, text="방향").grid(row=1, column=2, sticky="w", padx=4, pady=4)
            dir_opts = route.get("directions", []) + ["양방향"]
            self._create_styled_combobox(edit_body, variable=var_dir, values=dir_opts, state="readonly", width=100).grid(row=1, column=3, sticky="w", padx=4, pady=4)

            ctk.CTkLabel(edit_body, text="시작 이정(km)").grid(row=2, column=0, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(edit_body, textvariable=var_start, width=80).grid(row=2, column=1, sticky="w", padx=4, pady=4)
            ctk.CTkLabel(edit_body, text="교장(m)").grid(row=2, column=2, sticky="w", padx=4, pady=4)
            ctk.CTkEntry(edit_body, textvariable=var_length, width=80).grid(row=2, column=3, sticky="w", padx=4, pady=4)

            def on_save():
                try:
                    name = var_name.get().strip()
                    st_type = var_type.get()
                    direction = var_dir.get().strip()
                    start = float(var_start.get())
                    length = float(var_length.get())
                except ValueError:
                    self._show_error("입력 오류", "숫자를 올바르게 입력해주세요.")
                    return
                if not name or length <= 0:
                    self._show_error("입력 오류", "정보를 올바르게 입력해주세요.")
                    return

                end = start + length / 1000.0
                new_struct = {
                    'name': name, 'type': st_type, 'direction': direction,
                    'start': start, 'end': end, 'length': length
                }
                if is_edit:
                    route["structures"][struct_idx] = new_struct
                else:
                    route["structures"].append(new_struct)

                refresh_tree()
                self.draw_schematic()
                edit_dlg.destroy()

            btn_frame = ctk.CTkFrame(edit_dlg, fg_color="transparent")
            btn_frame.pack(fill="x", padx=10, pady=10)
            self._create_button(btn_frame, text="저장", command=on_save, width=80).pack(side="right")
            self._create_button(btn_frame, text="취소", command=edit_dlg.destroy, width=80, fg_color="transparent", border_width=1).pack(side="right", padx=5)

        def on_add():
            open_edit_dialog()

        def on_edit():
            selected = tree.selection()
            if not selected or selected[0] not in item_to_idx: return
            open_edit_dialog(item_to_idx[selected[0]])

        def on_delete():
            selected = tree.selection()
            if not selected or selected[0] not in item_to_idx: return
            if self._ask_yes_no("삭제 확인", "선택한 구조물을 삭제하시겠습니까?"):
                del route["structures"][item_to_idx[selected[0]]]
                refresh_tree()
                self.draw_schematic()

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)
        self._create_button(btn_frame, text="추가", command=on_add, width=80).pack(side="left")
        self._create_button(btn_frame, text="수정", command=on_edit, width=80).pack(side="left", padx=5)
        self._create_button(btn_frame, text="삭제", command=on_delete, width=80, fg_color="#C53030", hover_color="#9B2C2C").pack(side="left")
        self._create_close_button(btn_frame, dlg.destroy, width=80).pack(side="right")

    # ---------- 포장상태 데이터 관리 (제네릭 다이얼로그) ----------
    def _open_condition_data_dialog(self, data_key, title, value_type=float,
                                    redraw_on_save=True, csv_enabled=False):
        """DI/HPCI/AAR/RD/IRI 등 연도별 조건 데이터 관리 공용 다이얼로그."""
        if not (0 <= self.current_route_index < len(self.routes)):
            self._show_error("노선 없음", f"{title}을(를) 관리할 노선을 선택해 주세요.")
            return

        route = self.routes[self.current_route_index]
        if data_key not in route:
            route[data_key] = {}

        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry("1000x600")

        top_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        top_frame.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(top_frame, text=f"연도별 {title} 관리 (더블클릭하여 수정)").pack(side="left")

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=10)

        # 현재 연도 목록 수집
        current_years = set()
        for v_map in route[data_key].values():
            if isinstance(v_map, dict):
                current_years.update(v_map.keys())
            else:
                current_years.add(datetime.now().strftime("%Y"))
        current_years = sorted(current_years)
        if not current_years:
            current_years = [datetime.now().strftime("%Y")]

        tree_container = ctk.CTkFrame(body, fg_color="transparent")
        tree_container.pack(fill="both", expand=True)

        # tree 참조를 클로저에서 수정하기 위해 리스트로 래핑
        tree_ref = [None]

        def add_year():
            if len(current_years) >= 5:
                self._show_warning("경고", "최근 5개년만 기입 가능합니다.\n기존 데이터를 지워주세요.")
                return
            y = self._ask_string("연도 추가", "추가할 연도를 입력하세요 (예: 2025):")
            if y and y.isdigit() and len(y) == 4:
                if y not in current_years:
                    current_years.append(y)
                    current_years.sort()
                    refresh_tree()
            elif y:
                self._show_error("오류", "올바른 연도 형식이 아닙니다.")

        def delete_year():
            if not current_years:
                self._show_info("알림", "삭제할 연도가 없습니다.")
                return
            d_del = self._create_popup_window(dlg)
            d_del.title("")
            d_del.geometry("250x130")
            d_del.transient(dlg)
            d_del.grab_set()
            ctk.CTkLabel(d_del, text="삭제할 연도를 선택하세요:").pack(pady=10)
            var_del = tk.StringVar(value=current_years[0])
            self._create_styled_combobox(d_del, variable=var_del, values=current_years, state="readonly").pack(pady=5)
            def confirm_del():
                y_to_del = var_del.get()
                if self._ask_yes_no("삭제 확인", f"{y_to_del}년 데이터를 모두 삭제하시겠습니까?"):
                    if y_to_del in current_years:
                        current_years.remove(y_to_del)
                    for val_map in route[data_key].values():
                        if isinstance(val_map, dict) and y_to_del in val_map:
                            del val_map[y_to_del]
                    refresh_tree()
                    d_del.destroy()
            self._create_button(d_del, text="삭제", command=confirm_del).pack(pady=10)

        self._create_button(top_frame, text="연도 삭제", command=delete_year, width=80,
                      fg_color="#C53030", hover_color="#9B2C2C").pack(side="right", padx=5)
        self._create_button(top_frame, text="연도 추가", command=add_year, width=80).pack(side="right")

        r_start = route['start_km']
        r_end = route['end_km']
        directions = route.get('directions', [])
        lanes = [f"{i}차로" for i in range(1, 5)]

        def on_tree_double_click(event):
            tree = tree_ref[0]
            item_id = tree.identify_row(event.y)
            column_id = tree.identify_column(event.x)
            if not item_id or not column_id:
                return
            col_idx = int(column_id.replace('#', '')) - 1
            if col_idx < 4:
                return
            x, y, width, height = tree.bbox(item_id, column_id)
            curr_val = tree.item(item_id, "values")[col_idx]
            entry_var = tk.StringVar(value=curr_val)
            entry = ttk.Entry(tree, textvariable=entry_var)
            entry.place(x=x, y=y, width=width, height=height)
            entry.focus_set()
            def on_entry_save(event=None):
                new_value = entry_var.get()
                try:
                    if new_value:
                        value_type(new_value)
                except ValueError:
                    entry.destroy()
                    return
                current_values = list(tree.item(item_id, "values"))
                current_values[col_idx] = new_value
                tree.item(item_id, values=tuple(current_values))
                entry.destroy()
            entry.bind("<FocusOut>", on_entry_save)
            entry.bind("<Return>", on_entry_save)

        def refresh_tree():
            if tree_ref[0]:
                tree_ref[0].destroy()
            fixed_cols = ["direction", "lane", "start_km", "end_km"]
            fixed_headers = ["방향", "차로", "시작", "끝"]
            cols = fixed_cols + current_years
            tree = ttk.Treeview(tree_container, columns=cols, show="headings")
            tree_ref[0] = tree
            for c, h in zip(fixed_cols, fixed_headers):
                tree.heading(c, text=h)
                tree.column(c, width=60, anchor="center")
            for y in current_years:
                tree.heading(y, text=y)
                tree.column(y, width=60, anchor="center")
            vsb = ctk.CTkScrollbar(tree_container, orientation="vertical", command=tree.yview)
            hsb = ctk.CTkScrollbar(tree_container, orientation="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            tree_container.grid_rowconfigure(0, weight=1)
            tree_container.grid_columnconfigure(0, weight=1)
            tree.bind("<Double-1>", on_tree_double_click)
            for d in directions:
                for l in lanes:
                    current_km = km_floor_to_grid(r_start, GRID_KM)
                    while current_km < r_end:
                        start_seg = current_km
                        end_seg = current_km + GRID_KM
                        key = (d, l, f"{start_seg:.2f}")
                        val_map = route[data_key].get(key, {})
                        row_vals = [d, l, f"{start_seg:.2f}", f"{end_seg:.2f}"]
                        for y in current_years:
                            if isinstance(val_map, dict):
                                row_vals.append(val_map.get(y, ""))
                            elif y == datetime.now().strftime("%Y"):
                                row_vals.append(val_map)
                            else:
                                row_vals.append("")
                        tree.insert("", "end", values=row_vals)
                        current_km += GRID_KM

        refresh_tree()

        def on_save():
            tree = tree_ref[0]
            for item_id in tree.get_children():
                values = tree.item(item_id, "values")
                key = (values[0], values[1], values[2])
                if key not in route[data_key]:
                    route[data_key][key] = {}
                if not isinstance(route[data_key][key], dict):
                    route[data_key][key] = {}
                for i, year in enumerate(current_years):
                    val_str = values[4 + i]
                    if val_str and str(val_str).strip():
                        try:
                            route[data_key][key][year] = value_type(val_str)
                        except ValueError:
                            pass
                    elif year in route[data_key][key]:
                        del route[data_key][key][year]
            self.update_year_filter_values()
            if redraw_on_save:
                self.draw_schematic()
            dlg.destroy()

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)

        if csv_enabled:
            def on_export_csv():
                fpath = filedialog.asksaveasfilename(
                    title=f"{title} 데이터 CSV로 내보내기",
                    defaultextension=".csv",
                    filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
                    initialfile=f"{route['name']}_{data_key}.csv",
                    parent=dlg
                )
                if not fpath:
                    return
                try:
                    with open(fpath, "w", newline="", encoding="utf-8-sig") as f:
                        writer = csv.writer(f)
                        writer.writerow(["route_name", "direction", "lane", "start_km", "year", "value"])
                        for (d, l, skm), val_map in route[data_key].items():
                            if isinstance(val_map, dict):
                                for y, val in val_map.items():
                                    writer.writerow([route['name'], d, l, skm, y, val])
                            else:
                                writer.writerow([route['name'], d, l, skm, datetime.now().strftime("%Y"), val_map])
                    self._show_info("성공", f"{title} 데이터를 CSV 파일로 저장했습니다.")
                except Exception as ex:
                    self._show_error("오류", f"파일 저장 중 오류가 발생했습니다:\n{ex}")

            def on_import_csv():
                fpath = filedialog.askopenfilename(
                    title=f"{title} 데이터 CSV에서 가져오기",
                    filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
                    parent=dlg
                )
                if not fpath:
                    return
                imported_data = {}
                try:
                    with open(fpath, "r", newline="", encoding="utf-8-sig") as f:
                        reader = csv.DictReader(f)
                        for row in reader:
                            if row.get('route_name') != route['name']:
                                self._show_warning("불일치",
                                    f"CSV 파일의 노선명 '{row.get('route_name')}'이(가) 현재 노선 '{route['name']}'과(와) 다릅니다.")
                                continue
                            try:
                                val_str = row.get('value') or row.get(f'{data_key[:-5]}_value', '')
                                if val_str and val_str.strip():
                                    key = (row['direction'], row['lane'], f"{float(row['start_km']):.2f}")
                                    year = row.get('year', datetime.now().strftime("%Y"))
                                    if key not in imported_data:
                                        imported_data[key] = {}
                                    imported_data[key][year] = value_type(val_str)
                            except (KeyError, ValueError):
                                continue
                except Exception as ex:
                    self._show_error("오류", f"파일 읽기 중 오류가 발생했습니다:\n{ex}")
                    return
                for key, val_map in imported_data.items():
                    if key not in route[data_key]:
                        route[data_key][key] = {}
                    if isinstance(route[data_key][key], dict):
                        route[data_key][key].update(val_map)
                    else:
                        old_val = route[data_key][key]
                        route[data_key][key] = {datetime.now().strftime("%Y"): old_val}
                        route[data_key][key].update(val_map)
                self.update_year_filter_values()
                refresh_tree()
                self._show_info("가져오기 완료", "데이터를 불러왔습니다.")

            self._create_button(btn_frame, text="CSV 가져오기", command=on_import_csv, width=100).pack(side="left")
            self._create_button(btn_frame, text="CSV 내보내기", command=on_export_csv, width=100).pack(side="left", padx=5)

        self._create_button(btn_frame, text="저장", command=on_save, width=80).pack(side="right")
        self._create_close_button(btn_frame, dlg.destroy, width=80).pack(side="right", padx=5)

