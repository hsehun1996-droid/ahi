# -*- coding: utf-8 -*-
"""현황관리 - 통계 센터 (연도별 보수연장, 포장상태 추이)"""
import os
import sys
import math
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from constants import (
    PRIMARY_BLUE, PRIMARY_BLUE_HOVER, APP_BG, CARD_BG, CARD_BORDER,
    TITLE_TEXT, ACCENT_RED, DIRECTIONS, GRID_KM,
)

_CHART_COLORS = [
    "#2B6CB0", "#C53030", "#2F855A", "#D69E2E", "#805AD5",
    "#DD6B20", "#319795", "#D53F8C", "#4A5568", "#E53E3E",
    "#1A365D", "#742A2A", "#1C4532", "#744210", "#44337A",
]


class StatsMixin:

    def on_stats_dashboard(self):
        """현황관리 통계 센터 창을 엽니다."""
        if hasattr(self, '_stats_window') and self._stats_window and self._stats_window.winfo_exists():
            self._stats_window.lift()
            return

        win = self._create_popup_window(self)
        self._stats_window = win
        win.title("현황관리 — 통계 센터")
        win.geometry("1140x740")
        win.resizable(True, True)

        hdr = ctk.CTkFrame(win, fg_color=PRIMARY_BLUE, corner_radius=0, height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr, text="현황관리  |  통계 센터",
            font=(self.font_family, 16, "bold"), text_color="#FFFFFF",
        ).pack(side="left", padx=20)
        self._create_button(
            hdr, text="✕ 닫기", command=win.destroy,
            fg_color="transparent", text_color="#FFFFFF",
            hover_color=PRIMARY_BLUE_HOVER, width=80, height=32,
        ).pack(side="right", padx=10)

        tabs = ctk.CTkTabview(win, fg_color=APP_BG)
        tabs.pack(fill="both", expand=True)

        t1 = tabs.add("연도별 보수연장")
        t2 = tabs.add("포장상태 추이")

        self._stats_bar_rects = []
        self._stats_build_length_tab(t1)
        self._stats_build_condition_tab(t2)

    # ── Tab 1: 연도별 보수연장 ─────────────────────────────────────────────

    def _stats_build_length_tab(self, parent):
        ctrl = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=10,
                            border_width=1, border_color=CARD_BORDER)
        ctrl.pack(fill="x", padx=12, pady=(10, 6))

        ctk.CTkLabel(ctrl, text="그룹:", font=(self.font_family, 13, "bold"),
                     text_color=TITLE_TEXT).pack(side="left", padx=(14, 4), pady=10)

        self._stats_length_group = tk.StringVar(value="method")
        for val, label in [("method", "공법별"), ("hq", "본부별"), ("route", "노선별")]:
            ctk.CTkRadioButton(
                ctrl, text=label, variable=self._stats_length_group, value=val,
                command=self._stats_refresh_length,
                font=(self.font_family, 12), radiobutton_width=16, radiobutton_height=16,
            ).pack(side="left", padx=8, pady=10)

        ctk.CTkFrame(ctrl, width=1, height=24, fg_color=CARD_BORDER).pack(side="left", padx=10, pady=10)

        self._stats_include_plan = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            ctrl, text="사업계획 포함", variable=self._stats_include_plan,
            command=self._stats_refresh_length,
            font=(self.font_family, 12), checkbox_width=16, checkbox_height=16,
        ).pack(side="left", padx=6, pady=10)

        self._create_button(
            ctrl, text="↻ 새로고침", command=self._stats_refresh_length,
            width=100, height=30,
        ).pack(side="right", padx=12, pady=8)

        body = ctk.CTkFrame(parent, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        chart_frame = ctk.CTkFrame(body, fg_color=CARD_BG, corner_radius=10,
                                   border_width=1, border_color=CARD_BORDER)
        chart_frame.pack(fill="both", expand=True)

        self._stats_bar_canvas = tk.Canvas(chart_frame, bg="#FFFFFF",
                                           highlightthickness=0, height=310)
        self._stats_bar_canvas.pack(fill="both", expand=True, padx=2, pady=2)
        self._stats_bar_canvas.bind("<Configure>", lambda e: self._stats_refresh_length())
        self._stats_bar_canvas.bind("<Button-1>", self._stats_on_bar_click)

        self._stats_legend_frame = ctk.CTkFrame(body, fg_color="transparent")
        self._stats_legend_frame.pack(fill="x", pady=(4, 0))

        tbl_frame = ctk.CTkFrame(body, fg_color=CARD_BG, corner_radius=10,
                                  border_width=1, border_color=CARD_BORDER)
        tbl_frame.pack(fill="x", pady=(6, 0))
        self._stats_length_tree_frame = tbl_frame

        self._stats_refresh_length()

    def _stats_collect_length_data(self):
        group_by = self._stats_length_group.get()
        include_plan = self._stats_include_plan.get()

        data = {}
        for route in self.routes:
            route_name = route["name"]
            hq = (route.get("hq") or "").strip() or "미지정"
            for entry in route.get("entries", []):
                if not include_plan and entry.get("plan"):
                    continue
                wd = str(entry.get("work_date", "") or "")
                if len(wd) < 4 or not wd[:4].isdigit():
                    continue
                year = wd[:4]
                try:
                    length = abs(float(entry.get("end", 0) or 0) -
                                 float(entry.get("start", 0) or 0))
                except (TypeError, ValueError):
                    continue
                if length <= 0:
                    continue

                if group_by == "method":
                    cat = (entry.get("method") or "기타").strip() or "기타"
                elif group_by == "hq":
                    cat = hq
                else:
                    cat = route_name

                data.setdefault(year, {}).setdefault(cat, 0)
                data[year][cat] += length

        return data

    def _stats_draw_bar_chart(self, canvas, data):
        canvas.delete("all")
        self._stats_bar_rects = []

        W = canvas.winfo_width() or 900
        H = canvas.winfo_height() or 310

        if not data:
            canvas.create_text(W // 2, H // 2,
                                text="표시할 데이터가 없습니다.\n이력 입력 후 새로고침하세요.",
                                fill="#999999", font=(self.font_family, 12), justify="center")
            return

        ML, MR, MT, MB = 76, 24, 32, 58
        cw = W - ML - MR
        ch = H - MT - MB

        years = sorted(data.keys())

        all_cats = []
        for y in years:
            for c in data[y]:
                if c not in all_cats:
                    all_cats.append(c)

        cat_color = {c: _CHART_COLORS[i % len(_CHART_COLORS)] for i, c in enumerate(all_cats)}

        max_total = max(sum(data[y].values()) for y in years) if years else 1
        max_y = math.ceil(max_total * 1.15 * 10) / 10
        if max_y <= 0:
            max_y = 1

        y_steps = 5
        for i in range(y_steps + 1):
            val = max_y * i / y_steps
            py = MT + ch - (ch * val / max_y)
            canvas.create_line(ML, py, ML + cw, py, fill="#E8EEF5", width=1)
            canvas.create_text(ML - 6, py, text=f"{val:.2f}", anchor="e",
                                fill="#555555", font=(self.font_family, 9))

        canvas.create_text(
            16, MT + ch / 2, text="보수연장(km)", angle=90,
            fill="#444444", font=(self.font_family, 10, "bold"), anchor="center",
        )

        canvas.create_line(ML, MT, ML, MT + ch, fill="#333333", width=1.5)
        canvas.create_line(ML, MT + ch, ML + cw, MT + ch, fill="#333333", width=1.5)

        n = len(years)
        bar_w = min(72, cw / max(n + 1, 2) * 0.72)
        gap = cw / max(n + 1, 2)

        for i, year in enumerate(years):
            x_center = ML + gap * (i + 1)
            x0 = x_center - bar_w / 2
            year_total = sum(data[year].values())
            y_cursor = MT + ch

            for cat in all_cats:
                length = data[year].get(cat, 0)
                if length <= 0:
                    continue
                seg_h = ch * length / max_y
                y1 = y_cursor
                y0 = y_cursor - seg_h
                color = cat_color[cat]
                canvas.create_rectangle(x0, y0, x0 + bar_w, y1,
                                         fill=color, outline="#FFFFFF", width=1)
                self._stats_bar_rects.append(
                    (x0, y0, x0 + bar_w, y1, year, cat, length)
                )
                y_cursor = y0

            if year_total > 0:
                canvas.create_text(
                    x_center, y_cursor - 4, text=f"{year_total:.3f}",
                    fill="#222222", font=(self.font_family, 8, "bold"), anchor="s",
                )

            canvas.create_text(x_center, MT + ch + 16, text=year,
                                fill="#333333", font=(self.font_family, 10))

        self._stats_rebuild_legend(cat_color)
        self._stats_rebuild_table(data, years, all_cats)

    def _stats_rebuild_legend(self, cat_color):
        frame = self._stats_legend_frame
        for w in frame.winfo_children():
            w.destroy()
        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack()
        for cat, color in cat_color.items():
            item = ctk.CTkFrame(inner, fg_color="transparent")
            item.pack(side="left", padx=8, pady=2)
            box = tk.Canvas(item, width=14, height=14, bg=color,
                             highlightthickness=1, highlightbackground="#CCCCCC")
            box.pack(side="left", padx=(0, 3))
            ctk.CTkLabel(item, text=cat, font=(self.font_family, 11),
                         text_color=TITLE_TEXT).pack(side="left")

    def _stats_rebuild_table(self, data, years, all_cats):
        frame = self._stats_length_tree_frame
        for w in frame.winfo_children():
            w.destroy()

        cols = ["year"] + all_cats + ["total"]
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=5)
        tree.heading("year", text="연도")
        tree.column("year", width=60, anchor="center")
        col_w = max(70, min(150, 600 // max(len(all_cats), 1)))
        for c in all_cats:
            tree.heading(c, text=c)
            tree.column(c, width=col_w, anchor="center")
        tree.heading("total", text="합계(km)")
        tree.column("total", width=90, anchor="center")

        for year in sorted(years):
            vals = [year]
            total = 0.0
            for cat in all_cats:
                v = data[year].get(cat, 0)
                total += v
                vals.append(f"{v:.3f}" if v > 0 else "-")
            vals.append(f"{total:.3f}")
            tree.insert("", "end", values=vals)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

    def _stats_on_bar_click(self, event):
        x, y = event.x, event.y
        for (x0, y0, x1, y1, year, cat, length) in self._stats_bar_rects:
            if x0 <= x <= x1 and y0 <= y <= y1:
                self._stats_show_tooltip(event, f"{year}년 / {cat}\n{length:.3f} km")
                return

    def _stats_show_tooltip(self, event, text):
        canvas = self._stats_bar_canvas
        canvas.delete("tooltip")
        x, y = event.x, event.y
        pad = 6
        lbl = canvas.create_text(
            x + 14, y - 14, text=text, anchor="sw",
            fill="#1A202C", font=(self.font_family, 10, "bold"), tags="tooltip",
        )
        bb = canvas.bbox(lbl)
        if bb:
            canvas.create_rectangle(
                bb[0] - pad, bb[1] - pad, bb[2] + pad, bb[3] + pad,
                fill="#FFFFCC", outline="#AAAAAA", width=1, tags="tooltip",
            )
            canvas.tag_raise(lbl)
        canvas.after(2500, lambda: canvas.delete("tooltip"))

    def _stats_refresh_length(self):
        data = self._stats_collect_length_data()
        canvas = self._stats_bar_canvas
        canvas.update_idletasks()
        self._stats_draw_bar_chart(canvas, data)

    # ── Tab 2: 포장상태 추이 ──────────────────────────────────────────────

    def _stats_build_condition_tab(self, parent):
        ctrl = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=10,
                            border_width=1, border_color=CARD_BORDER)
        ctrl.pack(fill="x", padx=12, pady=(10, 6))

        ctk.CTkLabel(ctrl, text="지표:", font=(self.font_family, 13, "bold"),
                     text_color=TITLE_TEXT).pack(side="left", padx=(14, 4), pady=10)

        self._stats_cond_indicator = tk.StringVar(value="di")
        for val, label in [("di", "DI지수"), ("iri", "IRI등급"), ("hpci", "HPCI등급"), ("rd", "RD등급")]:
            ctk.CTkRadioButton(
                ctrl, text=label, variable=self._stats_cond_indicator, value=val,
                command=self._stats_refresh_condition,
                font=(self.font_family, 12), radiobutton_width=16, radiobutton_height=16,
            ).pack(side="left", padx=8, pady=10)

        ctk.CTkFrame(ctrl, width=1, height=24, fg_color=CARD_BORDER).pack(side="left", padx=10, pady=10)

        ctk.CTkLabel(ctrl, text="노선:", font=(self.font_family, 13, "bold"),
                     text_color=TITLE_TEXT).pack(side="left", padx=(6, 4), pady=10)

        route_names = ["전체"] + [r["name"] for r in self.routes]
        self._stats_cond_route = tk.StringVar(value="전체")
        self._stats_cond_route_cb = self._create_styled_combobox(
            ctrl, variable=self._stats_cond_route,
            values=route_names, width=160, state="readonly",
        )
        self._stats_cond_route_cb.pack(side="left", padx=4, pady=10)
        self._stats_cond_route_cb.bind(
            "<<ComboboxSelected>>", lambda e: self._stats_refresh_condition()
        )

        self._create_button(
            ctrl, text="↻ 새로고침", command=self._stats_refresh_condition,
            width=100, height=30,
        ).pack(side="right", padx=12, pady=8)

        chart_frame = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=10,
                                   border_width=1, border_color=CARD_BORDER)
        chart_frame.pack(fill="both", expand=True, padx=12, pady=(0, 4))

        self._stats_line_canvas = tk.Canvas(chart_frame, bg="#FFFFFF", highlightthickness=0)
        self._stats_line_canvas.pack(fill="both", expand=True, padx=2, pady=2)
        self._stats_line_canvas.bind("<Configure>", lambda e: self._stats_refresh_condition())

        self._stats_cond_legend_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._stats_cond_legend_frame.pack(fill="x", padx=12, pady=(2, 8))

        self._stats_refresh_condition()

    def _stats_collect_condition_data(self):
        indicator = self._stats_cond_indicator.get()
        route_filter = self._stats_cond_route.get()

        key_map = {
            "di": "di_data", "iri": "iri_data",
            "hpci": "hpci_data", "rd": "rd_data",
        }
        data_key = key_map.get(indicator, "di_data")
        # DI·HPCI: 값 ≥ 5인 셀의 연장(km) 합계 / IRI·RD: 구간 평균
        length_mode = indicator in ("di", "hpci")

        result = {}
        for route in self.routes:
            route_name = route["name"]
            if route_filter != "전체" and route_name != route_filter:
                continue

            data_map = route.get(data_key, {})

            if length_mode:
                year_lengths = {}
                for _cell_key, year_dict in data_map.items():
                    if not isinstance(year_dict, dict):
                        continue
                    for yr, val in year_dict.items():
                        try:
                            if float(val) >= 5.0:
                                year_lengths[yr] = year_lengths.get(yr, 0.0) + GRID_KM
                        except (ValueError, TypeError):
                            pass
                if year_lengths:
                    result[route_name] = year_lengths
            else:
                year_values = {}
                for _cell_key, year_dict in data_map.items():
                    if not isinstance(year_dict, dict):
                        continue
                    for yr, val in year_dict.items():
                        try:
                            year_values.setdefault(yr, []).append(float(val))
                        except (ValueError, TypeError):
                            pass
                if year_values:
                    result[route_name] = {
                        yr: sum(vs) / len(vs)
                        for yr, vs in year_values.items() if vs
                    }

        return result

    def _stats_draw_line_chart(self, canvas, data):
        canvas.delete("all")

        indicator = self._stats_cond_indicator.get()
        label_map = {
            "di":   "DI≥5 연장(km)",
            "hpci": "HPCI≥5 연장(km)",
            "iri":  "IRI등급 (평균)",
            "rd":   "RD등급 (평균)",
        }
        y_label = label_map.get(indicator, "지표값")

        W = canvas.winfo_width() or 900
        H = canvas.winfo_height() or 380

        if not data:
            canvas.create_text(
                W // 2, H // 2,
                text="선택한 노선의 포장상태 데이터가 없습니다.\n'포장상태' 메뉴에서 데이터를 먼저 입력해 주세요.",
                fill="#999999", font=(self.font_family, 12), justify="center",
            )
            return

        ML, MR, MT, MB = 76, 36, 36, 60
        cw = W - ML - MR
        ch = H - MT - MB

        all_years = sorted({yr for rd in data.values() for yr in rd})
        if not all_years:
            return

        route_names = list(data.keys())
        route_colors = {
            r: _CHART_COLORS[i % len(_CHART_COLORS)]
            for i, r in enumerate(route_names)
        }

        all_vals = [v for rd in data.values() for v in rd.values()]
        max_val = max(all_vals) if all_vals else 10
        length_mode = indicator in ("di", "hpci")
        if length_mode:
            y_min = 0
            y_max = max_val * 1.15 if max_val > 0 else 1
        else:
            min_val = min(all_vals) if all_vals else 0
            val_range = max_val - min_val
            if val_range < 0.001:
                val_range = 1
            y_min = max(0, min_val - val_range * 0.15)
            y_max = max_val + val_range * 0.15
        y_range = y_max - y_min if (y_max - y_min) > 0 else 1

        y_steps = 5
        for i in range(y_steps + 1):
            v = y_min + y_range * i / y_steps
            py = MT + ch - (ch * (v - y_min) / y_range)
            canvas.create_line(ML, py, ML + cw, py, fill="#E8EEF5", width=1)
            canvas.create_text(ML - 6, py, text=f"{v:.2f}", anchor="e",
                                fill="#555555", font=(self.font_family, 9))

        canvas.create_text(
            16, MT + ch / 2, text=y_label, angle=90,
            fill="#444444", font=(self.font_family, 10, "bold"), anchor="center",
        )
        canvas.create_line(ML, MT, ML, MT + ch, fill="#333333", width=1.5)
        canvas.create_line(ML, MT + ch, ML + cw, MT + ch, fill="#333333", width=1.5)

        n = len(all_years)
        if n == 1:
            x_positions = {all_years[0]: ML + cw / 2}
        else:
            x_positions = {
                yr: ML + cw * i / (n - 1)
                for i, yr in enumerate(all_years)
            }

        for yr, px in x_positions.items():
            canvas.create_text(px, MT + ch + 16, text=yr,
                                fill="#333333", font=(self.font_family, 10))

        for route_name, yr_data in data.items():
            color = route_colors[route_name]
            pts = []
            for yr in all_years:
                if yr not in yr_data:
                    continue
                px = x_positions[yr]
                v = yr_data[yr]
                py = MT + ch - (ch * (v - y_min) / y_range)
                pts.append((px, py))
                canvas.create_oval(px - 5, py - 5, px + 5, py + 5,
                                    fill=color, outline="#FFFFFF", width=2)
                canvas.create_text(px, py - 12, text=f"{v:.2f}",
                                    fill=color, font=(self.font_family, 9, "bold"), anchor="s")

            if len(pts) >= 2:
                flat = []
                for px, py in pts:
                    flat += [px, py]
                canvas.create_line(*flat, fill=color, width=2.5, smooth=False)

        self._stats_rebuild_cond_legend(route_colors)

    def _stats_rebuild_cond_legend(self, route_colors):
        frame = self._stats_cond_legend_frame
        for w in frame.winfo_children():
            w.destroy()
        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack()
        for route_name, color in route_colors.items():
            item = ctk.CTkFrame(inner, fg_color="transparent")
            item.pack(side="left", padx=10, pady=2)
            line_c = tk.Canvas(item, width=28, height=4, bg=color, highlightthickness=0)
            line_c.pack(side="left", padx=(0, 4))
            ctk.CTkLabel(item, text=route_name, font=(self.font_family, 11),
                         text_color=TITLE_TEXT).pack(side="left")

    def _stats_refresh_condition(self):
        route_names = ["전체"] + [r["name"] for r in self.routes]
        try:
            self._stats_cond_route_cb.configure(values=route_names)
        except Exception:
            pass
        data = self._stats_collect_condition_data()
        canvas = self._stats_line_canvas
        canvas.update_idletasks()
        self._stats_draw_line_chart(canvas, data)
