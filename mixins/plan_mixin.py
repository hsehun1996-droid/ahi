# -*- coding: utf-8 -*-
"""사업계획 / 운영계획변경 관리 믹스인.

워크플로
─────────
1) 사업계획(on_business_plan)
   - '개량 우선순위 산정'으로 후보를 산정·다중선택 → '사업계획 적용'으로 표에 추가
   - 표에서 이정/방향/차로/공법 수정(공법은 범례 공법만 드롭다운)
   - '사업계획 확정' → 보수이력에 계획 엔트리(plan=True)로 주입 → 모식도에 '계획' 표기
2) 운영계획변경(on_operation_plan_change)
   - 후보 = (보수 필요 구간) - (사업계획 지정 구간)
   - 노선별 보기. 사업계획 구간은 이미 표출(비고=사업계획), 신규 선택분(비고=운영계획 변경)
   - '한글 내보내기' → 단위사업명/목적/단가 입력 → HWP 양식(operation_plan_template.hwp) 작성
"""
import os
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
from datetime import datetime

from constants import (
    PRIMARY_BLUE, PRIMARY_BLUE_HOVER, TITLE_TEXT, METHOD_STYLES,
)
from utils import (
    get_user_data_dir, log_exception, read_csv_with_encoding,
)

BUSINESS_PLAN_CSV = "all_business_plan.csv"
OPERATION_CHANGE_CSV = "all_operation_change.csv"

DEFAULT_PURPOSE = (
    "포장노후화 및 동절기 내 제설작업에 의한 노면파손이 발생함에 따라 "
    "안전한 주행환경 제공을 위해 포장개량 추진"
)


def _classify_pavement(pav: str) -> str:
    """포장형식 문자열을 '콘크리트' / '아스팔트'로 분류."""
    s = str(pav or "")
    if "아스팔트" in s or s.upper().startswith("A"):
        return "아스팔트"
    return "콘크리트"  # 기본값(콘크리트)


def _round_amount(value: float) -> int:
    """사업비(백만원)를 10단위로 반올림(0.5는 올림)."""
    import math
    try:
        return int(math.floor(float(value) / 10.0 + 0.5)) * 10
    except Exception:
        return 0


def _fmt_won(v) -> str:
    try:
        return f"{int(round(float(v))):,}"
    except Exception:
        return str(v)


class PlanMixin:
    # ───────────────────────────── 공용 편집 테이블 ─────────────────────────────
    def _build_plan_table(self, parent, columns, data_list,
                          on_data_change=None, row_filter=None, tag_func=None):
        """더블클릭 인라인 편집이 가능한 Treeview를 만든다.

        columns : [{key,label,width,kind}]  kind ∈ {text,num,method,readonly}
        data_list : 행(dict) 목록 (참조로 직접 갱신)
        on_data_change(row_idx, key) : 값 변경 후 콜백
        row_filter(row) -> bool : False면 표시 제외 (iid는 data_list 인덱스 유지)
        tag_func(row) -> tag명 : 행 태그(색상 등)
        반환: (tree, refresh)
        """
        style = ttk.Style(parent)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Plan.Treeview.Heading", background="#4A5568",
                        foreground="white", relief="flat",
                        font=(self.font_family, 10, "bold"))
        style.configure("Plan.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF",
                        foreground="#1A202C", rowheight=26, borderwidth=0)
        style.map("Plan.Treeview", background=[("selected", "#BEE3F8")],
                  foreground=[("selected", "#1A202C")])

        col_keys = [c["key"] for c in columns]
        tree = ttk.Treeview(parent, columns=col_keys, show="headings",
                            selectmode="browse", style="Plan.Treeview", height=12)
        for c in columns:
            tree.heading(c["key"], text=c["label"])
            tree.column(c["key"], width=c.get("width", 90), anchor="center")

        def _display(row, c):
            key = c["key"]
            if key == "length":
                try:
                    return f"{float(row.get('end', 0)) - float(row.get('start', 0)):.2f}"
                except Exception:
                    return ""
            v = row.get(key, "")
            if c.get("kind") == "num":
                try:
                    return f"{float(v):.2f}"
                except Exception:
                    return str(v)
            return str(v)

        def refresh():
            for iid in tree.get_children():
                tree.delete(iid)
            for i, row in enumerate(data_list):
                if row_filter and not row_filter(row):
                    continue
                tags = ()
                if tag_func:
                    t = tag_func(row)
                    if t:
                        tags = (t,)
                tree.insert("", "end", iid=str(i),
                            values=[_display(row, c) for c in columns], tags=tags)

        editor = {"widget": None, "info": None}

        def _close_editor(save):
            w = editor["widget"]
            if w is None:
                return
            info = editor["info"]
            editor["widget"] = None
            editor["info"] = None
            if save and info is not None:
                try:
                    val = w.get()
                except Exception:
                    val = ""
                if 0 <= info["row"] < len(data_list):
                    row = data_list[info["row"]]
                    c = info["col"]
                    key = c["key"]
                    if c.get("kind") == "num":
                        try:
                            fv = float(str(val).replace(",", "").replace("k", "").strip())
                            row[key] = round(fv, 3)
                        except Exception:
                            pass
                    else:
                        row[key] = val
                    if on_data_change:
                        try:
                            on_data_change(info["row"], key)
                        except Exception:
                            log_exception("plan table on_data_change 실패")
            try:
                w.destroy()
            except Exception:
                pass
            refresh()

        def _begin_edit(event):
            if editor["widget"] is not None:
                _close_editor(True)
            if tree.identify_region(event.x, event.y) != "cell":
                return
            rowid = tree.identify_row(event.y)
            colid = tree.identify_column(event.x)
            if not rowid or not colid:
                return
            cidx = int(colid[1:]) - 1
            if cidx < 0 or cidx >= len(columns):
                return
            c = columns[cidx]
            if c.get("kind") == "readonly" or c["key"] == "length":
                return
            try:
                ridx = int(rowid)
            except ValueError:
                return
            bbox = tree.bbox(rowid, colid)
            if not bbox:
                return
            x, y, w, h = bbox
            row = data_list[ridx]
            if c.get("kind") == "method":
                wid = ttk.Combobox(tree, values=list(METHOD_STYLES.keys()), state="readonly")
                wid.set(str(row.get(c["key"], "")))
            else:
                wid = ttk.Entry(tree)
                wid.insert(0, _display(row, c))
            wid.place(x=x, y=y, width=w, height=h)
            wid.focus_set()
            editor["widget"] = wid
            editor["info"] = {"row": ridx, "col": c}
            wid.bind("<Return>", lambda e: _close_editor(True))
            wid.bind("<Escape>", lambda e: _close_editor(False))
            wid.bind("<FocusOut>", lambda e: _close_editor(True))
            if c.get("kind") == "method":
                wid.bind("<<ComboboxSelected>>", lambda e: _close_editor(True))

        tree.bind("<Double-1>", _begin_edit)
        refresh()
        return tree, refresh

    @staticmethod
    def _section_overlap(a, b):
        """동일 노선/방향/차로 + 이정 겹침 여부."""
        if str(a.get("route")) != str(b.get("route")):
            return False
        if str(a.get("direction")) != str(b.get("direction")):
            return False
        if str(a.get("lane")) != str(b.get("lane")):
            return False
        try:
            return (float(a.get("start", 0)) < float(b.get("end", 0)) and
                    float(a.get("end", 0)) > float(b.get("start", 0)))
        except Exception:
            return False

    # ───────────────────────────── 사업계획 ─────────────────────────────
    def on_business_plan(self):
        """사업계획 창: 개량 우선순위 산정 + 사업계획 작성/확정."""
        if (self.business_plan_window is not None
                and self.business_plan_window.winfo_exists()):
            self.business_plan_window.lift()
            self.business_plan_window.focus_force()
            return

        dlg = self._create_popup_window(self)
        self.business_plan_window = dlg
        dlg.title("")
        dlg.geometry("1040x620")
        dlg.transient(self)
        dlg.lift()
        dlg.focus_force()

        top = ctk.CTkFrame(dlg, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(12, 4))
        ctk.CTkLabel(top, text=f"{self.plan_year}년 사업계획",
                     font=(self.font_family, 17, "bold"),
                     text_color=TITLE_TEXT).pack(side="left")
        self._create_button(top, text="개량 우선순위 산정",
                            command=self._bp_open_priority, width=160,
                            fg_color=PRIMARY_BLUE, hover_color=PRIMARY_BLUE_HOVER,
                            text_color="#FFFFFF",
                            font=(self.font_family, 13, "bold")).pack(side="right", padx=4)

        ctk.CTkLabel(
            dlg,
            text="‘개량 우선순위 산정’으로 구간을 선택해 추가하고, 표에서 이정·방향·차로·공법을 "
                 "수정한 뒤 ‘사업계획 확정’을 누르세요. (셀 더블클릭 = 수정)",
            font=(self.font_family, 11), text_color="#5A6B82", anchor="w", justify="left",
        ).pack(fill="x", padx=16, pady=(0, 6))

        table_frame = ctk.CTkFrame(dlg, fg_color="#FFFFFF", corner_radius=8)
        table_frame.pack(fill="both", expand=True, padx=14, pady=6)
        columns = [
            {"key": "route", "label": "노선명", "width": 120, "kind": "text"},
            {"key": "direction", "label": "방향", "width": 100, "kind": "text"},
            {"key": "lane", "label": "차로", "width": 70, "kind": "text"},
            {"key": "start", "label": "시작(km)", "width": 85, "kind": "num"},
            {"key": "end", "label": "끝(km)", "width": 85, "kind": "num"},
            {"key": "length", "label": "연장(km)", "width": 85, "kind": "readonly"},
            {"key": "method", "label": "공법", "width": 160, "kind": "method"},
        ]
        tree, refresh = self._build_plan_table(table_frame, columns, self.business_plan)
        tree.pack(fill="both", expand=True, padx=6, pady=6)
        self._bp_tree = tree
        self._bp_refresh = refresh

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(fill="x", padx=14, pady=(2, 12))
        self._create_button(btns, text="선택 행 삭제",
                            command=lambda: self._bp_delete_selected(), width=110,
                            fg_color="#A0AEC0", hover_color="#8A98AB").pack(side="left", padx=3)
        self._create_button(btns, text="전체 비우기",
                            command=lambda: self._bp_clear(), width=110,
                            fg_color="#A0AEC0", hover_color="#8A98AB").pack(side="left", padx=3)
        self._create_button(btns, text="사업계획 확정", command=self._bp_confirm,
                            width=150, fg_color="#2F855A", hover_color="#276749",
                            text_color="#FFFFFF",
                            font=(self.font_family, 13, "bold")).pack(side="right", padx=3)

        def _on_close():
            self.business_plan_window = None
            try:
                dlg.destroy()
            except Exception:
                pass
        dlg.protocol("WM_DELETE_WINDOW", _on_close)

    def _bp_open_priority(self):
        self.on_improvement_priority(
            apply_label="사업계획 적용",
            apply_callback=self._bp_apply_priority,
            default_year=self.plan_year,
        )

    def _bp_apply_priority(self, items, plan_year):
        """우선순위 팝업에서 선택한 구간을 사업계획 표에 추가."""
        self.plan_year = str(plan_year)
        added = 0
        for it in items:
            row = {
                "route": it.get("route", ""),
                "direction": it.get("direction", ""),
                "lane": it.get("lane", "전차로"),
                "start": round(float(it.get("start", 0)), 3),
                "end": round(float(it.get("end", 0)), 3),
                "method": it.get("method", "절삭 덧씌우기"),
                "pavement": it.get("pavement", ""),
                "year": str(plan_year),
            }
            if any(self._section_overlap(row, ex) for ex in self.business_plan):
                continue
            self.business_plan.append(row)
            added += 1
        # 새 연도가 적용되었으니 창 제목 갱신을 위해 창을 다시 띄움
        if (self.business_plan_window is None
                or not self.business_plan_window.winfo_exists()):
            self.on_business_plan()
        else:
            self._bp_refresh()
            self.business_plan_window.lift()
            self.business_plan_window.focus_force()
        self._show_info("적용", f"{added}개 구간을 사업계획에 추가했습니다.")

    def _bp_delete_selected(self):
        tree = getattr(self, "_bp_tree", None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        idxs = sorted((int(i) for i in sel), reverse=True)
        for i in idxs:
            if 0 <= i < len(self.business_plan):
                del self.business_plan[i]
        self._bp_refresh()

    def _bp_clear(self):
        if not self.business_plan:
            return
        if not messagebox.askyesno("확인", "사업계획 표를 모두 비울까요?"):
            return
        self.business_plan = []
        self._bp_refresh()

    def _bp_confirm(self):
        """사업계획 확정 → 보수이력에 계획 엔트리 주입, 저장, 모식도 갱신."""
        if not self.business_plan:
            self._show_info("알림", "사업계획 구간이 없습니다.")
            return
        self._inject_plan_entries()
        self.business_plan_confirmed = True
        base_dir = get_user_data_dir()
        try:
            self._save_business_plan(base_dir)
        except Exception:
            log_exception("사업계획 저장 실패")
        for fn in ("update_year_filter_values", "draw_schematic", "refresh_dashboard"):
            try:
                getattr(self, fn)()
            except Exception:
                log_exception(f"{fn} 실패")
        self._show_info(
            "완료",
            f"{self.plan_year}년 사업계획이 확정되었습니다.\n모식도에 ‘계획’으로 표시됩니다.",
        )

    def _inject_plan_entries(self):
        """business_plan 항목을 노선별 entries에 plan=True 엔트리로 주입(중복 방지)."""
        for r in self.routes:
            r["entries"] = [e for e in r.get("entries", []) if not e.get("plan")]
        name_to_route = {r.get("name"): r for r in self.routes}
        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        for item in self.business_plan:
            r = name_to_route.get(item.get("route"))
            if r is None:
                continue
            yr = str(item.get("year") or self.plan_year)
            r.setdefault("entries", []).append({
                "start": float(item.get("start", 0)),
                "end": float(item.get("end", 0)),
                "method": item.get("method", "절삭 덧씌우기"),
                "ts": ts,
                "direction": item.get("direction", ""),
                "lane": item.get("lane", "전차로"),
                "work_date": f"{yr}0101",
                "plan": True,
            })
            try:
                self._mark_route_data_changed(r)
            except Exception:
                pass

    # ───────────────────────────── 운영계획변경 ─────────────────────────────
    def on_operation_plan_change(self):
        """운영계획변경 창: (보수필요-사업계획) 산정 + 사업계획 선표출 + HWP 내보내기."""
        if (self.operation_change_window is not None
                and self.operation_change_window.winfo_exists()):
            self.operation_change_window.lift()
            self.operation_change_window.focus_force()
            return

        self._oc_sync_from_business_plan()

        dlg = self._create_popup_window(self)
        self.operation_change_window = dlg
        dlg.title("")
        dlg.geometry("1120x650")
        dlg.transient(self)
        dlg.lift()
        dlg.focus_force()

        top = ctk.CTkFrame(dlg, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(12, 4))
        ctk.CTkLabel(top, text="운영계획변경",
                     font=(self.font_family, 17, "bold"),
                     text_color=TITLE_TEXT).pack(side="left")

        ctk.CTkLabel(top, text="노선:", font=(self.font_family, 12)).pack(side="left", padx=(16, 4))
        route_names = ["전체"] + [r.get("name", "") for r in self.routes]
        self._oc_route_var = tk.StringVar(value="전체")
        route_menu = self._create_styled_option_menu(
            top, variable=self._oc_route_var, values=route_names, width=140,
            command=lambda *_: self._oc_refresh())
        route_menu.pack(side="left")

        self._create_button(top, text="개량 우선순위 산정",
                            command=self._oc_open_priority, width=160,
                            fg_color=PRIMARY_BLUE, hover_color=PRIMARY_BLUE_HOVER,
                            text_color="#FFFFFF",
                            font=(self.font_family, 13, "bold")).pack(side="right", padx=4)
        self._create_button(top, text="한글 내보내기",
                            command=self._oc_export_hwp, width=140,
                            fg_color="#2F855A", hover_color="#276749",
                            text_color="#FFFFFF",
                            font=(self.font_family, 13, "bold")).pack(side="right", padx=4)

        ctk.CTkLabel(
            dlg,
            text="비고=‘사업계획’은 확정된 사업계획 구간(당초), ‘운영계획 변경’은 신규 추가 구간입니다. "
                 "이정·차로·방향·공법 수정 가능(셀 더블클릭). 사업계획 구간의 이정을 바꾸면 한글 양식에서 "
                 "당초→변경으로 반영됩니다.",
            font=(self.font_family, 11), text_color="#5A6B82", anchor="w", justify="left",
            wraplength=1080,
        ).pack(fill="x", padx=16, pady=(0, 6))

        table_frame = ctk.CTkFrame(dlg, fg_color="#FFFFFF", corner_radius=8)
        table_frame.pack(fill="both", expand=True, padx=14, pady=6)
        columns = [
            {"key": "route", "label": "노선명", "width": 110, "kind": "text"},
            {"key": "direction", "label": "방향", "width": 95, "kind": "text"},
            {"key": "lane", "label": "차로", "width": 65, "kind": "text"},
            {"key": "start", "label": "시작(km)", "width": 80, "kind": "num"},
            {"key": "end", "label": "끝(km)", "width": 80, "kind": "num"},
            {"key": "length", "label": "연장(km)", "width": 80, "kind": "readonly"},
            {"key": "method", "label": "공법", "width": 150, "kind": "method"},
            {"key": "note", "label": "비고", "width": 110, "kind": "readonly"},
        ]

        def _row_filter(row):
            sel = self._oc_route_var.get()
            return sel == "전체" or str(row.get("route")) == sel

        def _tag_func(row):
            return "plan_row" if row.get("note") == "사업계획" else "change_row"

        tree, refresh = self._build_plan_table(
            table_frame, columns, self.operation_changes,
            row_filter=_row_filter, tag_func=_tag_func)
        tree.tag_configure("plan_row", background="#EBF8FF")
        tree.tag_configure("change_row", background="#FFFFFF")
        tree.pack(fill="both", expand=True, padx=6, pady=6)
        self._oc_tree = tree
        self._oc_refresh = refresh
        refresh()

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(fill="x", padx=14, pady=(2, 12))
        self._create_button(btns, text="선택 행 삭제",
                            command=lambda: self._oc_delete_selected(), width=110,
                            fg_color="#A0AEC0", hover_color="#8A98AB").pack(side="left", padx=3)

        def _on_close():
            self.operation_change_window = None
            try:
                base_dir = get_user_data_dir()
                self._save_operation_changes(base_dir)
            except Exception:
                log_exception("운영계획변경 저장 실패")
            try:
                dlg.destroy()
            except Exception:
                pass
        dlg.protocol("WM_DELETE_WINDOW", _on_close)

    def _oc_sync_from_business_plan(self):
        """확정된 사업계획 구간을 운영계획변경 목록에 비고=사업계획으로 동기화(누락분만 추가)."""
        existing_plan = [r for r in self.operation_changes if r.get("note") == "사업계획"]
        for item in self.business_plan:
            # orig 좌표 기준으로 이미 있는지 확인
            already = any(
                str(p.get("route")) == str(item.get("route"))
                and str(p.get("direction")) == str(item.get("direction"))
                and str(p.get("lane")) == str(item.get("lane"))
                and abs(float(p.get("orig_start", -999)) - float(item.get("start", 0))) < 1e-6
                and abs(float(p.get("orig_end", -999)) - float(item.get("end", 0))) < 1e-6
                for p in existing_plan
            )
            if already:
                continue
            self.operation_changes.append({
                "route": item.get("route", ""),
                "direction": item.get("direction", ""),
                "lane": item.get("lane", "전차로"),
                "start": round(float(item.get("start", 0)), 3),
                "end": round(float(item.get("end", 0)), 3),
                "method": item.get("method", "절삭 덧씌우기"),
                "pavement": item.get("pavement", ""),
                "note": "사업계획",
                "orig_start": round(float(item.get("start", 0)), 3),
                "orig_end": round(float(item.get("end", 0)), 3),
            })

    def _oc_open_priority(self):
        # 후보에서 사업계획 지정 구간 제외
        exclude = [
            {"route": p.get("route"), "direction": p.get("direction"),
             "lane": p.get("lane"),
             "start": float(p.get("orig_start", p.get("start", 0))),
             "end": float(p.get("orig_end", p.get("end", 0)))}
            for p in self.operation_changes if p.get("note") == "사업계획"
        ]
        exclude += [
            {"route": p.get("route"), "direction": p.get("direction"),
             "lane": p.get("lane"), "start": float(p.get("start", 0)),
             "end": float(p.get("end", 0))}
            for p in self.business_plan
        ]
        self.on_improvement_priority(
            apply_label="운영계획변경 적용",
            apply_callback=self._oc_apply_priority,
            exclude_sections=exclude,
            default_year=self.plan_year,
        )

    def _oc_apply_priority(self, items, plan_year):
        added = 0
        for it in items:
            row = {
                "route": it.get("route", ""),
                "direction": it.get("direction", ""),
                "lane": it.get("lane", "전차로"),
                "start": round(float(it.get("start", 0)), 3),
                "end": round(float(it.get("end", 0)), 3),
                "method": it.get("method", "절삭 덧씌우기"),
                "pavement": it.get("pavement", ""),
                "note": "운영계획 변경",
                "orig_start": None,
                "orig_end": None,
            }
            if any(self._section_overlap(row, ex) for ex in self.operation_changes):
                continue
            self.operation_changes.append(row)
            added += 1
        if (self.operation_change_window is None
                or not self.operation_change_window.winfo_exists()):
            self.on_operation_plan_change()
        else:
            self._oc_refresh()
            self.operation_change_window.lift()
            self.operation_change_window.focus_force()
        self._show_info("적용", f"{added}개 구간을 운영계획변경에 추가했습니다.")

    def _oc_delete_selected(self):
        tree = getattr(self, "_oc_tree", None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        idxs = sorted((int(i) for i in sel), reverse=True)
        for i in idxs:
            if 0 <= i < len(self.operation_changes):
                del self.operation_changes[i]
        self._oc_refresh()

    # ───────────────────────────── HWP 내보내기 ─────────────────────────────
    def _oc_export_hwp(self):
        if not self.operation_changes:
            self._show_info("알림", "운영계획변경 구간이 없습니다.")
            return

        # 사업계획 구간 중 이정이 수정된 항목이 있는지 확인
        modified = any(
            r.get("note") == "사업계획" and (
                abs(float(r.get("start", 0)) - float(r.get("orig_start", r.get("start", 0)))) > 1e-6
                or abs(float(r.get("end", 0)) - float(r.get("orig_end", r.get("end", 0)))) > 1e-6
            )
            for r in self.operation_changes
        )
        if modified:
            if not messagebox.askyesno(
                "운영계획변경 적용",
                "사업계획 구간의 이정이 수정되었습니다.\n변경된 이정으로 운영계획변경을 적용하시겠습니까?"):
                return

        self._open_hwp_input_dialog()

    def _open_hwp_input_dialog(self):
        dlg = self._create_popup_window(self)
        dlg.title("")
        dlg.geometry("520x440")
        dlg.transient(self)
        dlg.grab_set()
        dlg.lift()
        dlg.focus_force()

        ctk.CTkLabel(dlg, text="운영계획변경 - 한글 양식 작성",
                     font=(self.font_family, 15, "bold"),
                     text_color=TITLE_TEXT).pack(anchor="w", padx=18, pady=(16, 10))

        frm = ctk.CTkFrame(dlg, fg_color="transparent")
        frm.pack(fill="both", expand=True, padx=18)

        ctk.CTkLabel(frm, text="단위사업명", font=(self.font_family, 12, "bold"),
                     text_color=TITLE_TEXT).pack(anchor="w")
        var_name = tk.StringVar(value="")
        ctk.CTkEntry(frm, textvariable=var_name, width=460,
                     placeholder_text="예) 중앙선 246.9-313.5k 콘크리트포장 덧씌우기").pack(anchor="w", pady=(2, 10))

        ctk.CTkLabel(frm, text="목적", font=(self.font_family, 12, "bold"),
                     text_color=TITLE_TEXT).pack(anchor="w")
        txt_purpose = ctk.CTkTextbox(frm, width=460, height=90,
                                     font=(self.font_family, 12), wrap="word")
        txt_purpose.pack(anchor="w", pady=(2, 10))
        txt_purpose.insert("end", DEFAULT_PURPOSE)

        ctk.CTkLabel(frm, text="단가 (백만원/km)", font=(self.font_family, 12, "bold"),
                     text_color=TITLE_TEXT).pack(anchor="w")
        var_price = tk.StringVar(value="")
        ctk.CTkEntry(frm, textvariable=var_price, width=200,
                     placeholder_text="예) 370").pack(anchor="w", pady=(2, 6))

        ctk.CTkLabel(frm,
                     text="사업비 = 연장(km) × 단가, 백만원 단위 10단위 반올림",
                     font=(self.font_family, 10), text_color="#5A6B82").pack(anchor="w")

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.pack(fill="x", padx=18, pady=14)

        def _do_export():
            name = var_name.get().strip()
            purpose = txt_purpose.get("1.0", "end").strip()
            try:
                unit_price = float(var_price.get().replace(",", "").strip())
            except Exception:
                self._show_error("입력 오류", "단가를 숫자로 입력해 주세요.")
                return
            if not name:
                self._show_error("입력 오류", "단위사업명을 입력해 주세요.")
                return
            try:
                dlg.destroy()
            except Exception:
                pass
            self._run_hwp_export(name, purpose, unit_price)

        self._create_button(btns, text="내보내기", command=_do_export, width=120,
                            fg_color="#2F855A", hover_color="#276749",
                            text_color="#FFFFFF",
                            font=(self.font_family, 13, "bold")).pack(side="right", padx=4)
        self._create_close_button(btns, dlg.destroy, width=100).pack(side="right", padx=4)

    def _build_hwp_payload(self, project_name, purpose, unit_price):
        """운영계획변경 데이터 → HWP 양식용 구조로 변환."""
        rows = []
        sum_orig = sum_new = sum_delta = 0
        sum_len = 0.0
        for r in self.operation_changes:
            cur_len = round(float(r.get("end", 0)) - float(r.get("start", 0)), 2)
            is_plan = (r.get("note") == "사업계획")
            if is_plan:
                orig_len = round(float(r.get("orig_end", r.get("end", 0)))
                                 - float(r.get("orig_start", r.get("start", 0))), 2)
            else:
                orig_len = 0.0
            cur_amt = _round_amount(cur_len * unit_price)
            orig_amt = _round_amount(orig_len * unit_price) if is_plan else None

            if is_plan:
                init_v = orig_amt
                change_v = cur_amt
                delta_v = change_v - init_v
            else:
                init_v = None          # 신규: 당초 없음
                change_v = cur_amt
                delta_v = cur_amt

            # 위치 표기: 이정 + (방향)
            direction = str(r.get("direction", "")).replace("방향", "").strip()
            loc = f"{float(r.get('start', 0)):.1f}~{float(r.get('end', 0)):.2f}k"
            if direction:
                loc += f"({direction})"
            lane = str(r.get("lane", "")).replace("차로", "").strip()

            rows.append({
                "route": str(r.get("route", "")).split()[0] if r.get("route") else "",
                "location": loc,
                "lane": lane,
                "length": f"{cur_len:.1f}",
                "init": _fmt_won(init_v) if init_v is not None else "-",
                "change": _fmt_won(change_v),
                "delta": _fmt_won(delta_v),
                "di": "5~7",
                "note": r.get("note", ""),
            })
            sum_len += cur_len
            sum_orig += (init_v or 0)
            sum_new += change_v
            sum_delta += delta_v

        totals = {
            "count": f"{len(rows)}개소",
            "length": f"{sum_len:.1f}",
            "init": _fmt_won(sum_orig),
            "change": _fmt_won(sum_new),
            "delta": _fmt_won(sum_delta),
        }

        # 소요예산 산출내역: 포장형식별 집계
        budget = {}
        for r in self.operation_changes:
            cls = _classify_pavement(r.get("pavement"))
            cur_len = round(float(r.get("end", 0)) - float(r.get("start", 0)), 2)
            budget.setdefault(cls, 0.0)
            budget[cls] += cur_len
        budget_rows = []
        for cls, total_len in budget.items():
            cost = _round_amount(total_len * unit_price)
            budget_rows.append({
                "content": f"{cls}포장 절삭 덧씌우기(t=10cm)",
                "qty": f"{total_len:.2f}",
                "unit_price": _fmt_won(unit_price),
                "cost": _fmt_won(cost),
                "note": "야간",
            })

        budget_change = {
            "name": project_name,
            "init": _fmt_won(sum_orig),
            "change": _fmt_won(sum_new),
            "delta": _fmt_won(sum_delta),
        }

        return {
            "project_name": project_name,
            "purpose": purpose,
            "unit_price": unit_price,
            "table1_rows": rows,
            "table1_totals": totals,
            "table2_rows": budget_rows,
            "table3": budget_change,
        }

    def _run_hwp_export(self, project_name, purpose, unit_price):
        try:
            import hwp_export
        except Exception:
            log_exception("hwp_export 모듈 로드 실패")
            self._show_error("오류", "한글 내보내기 모듈을 불러올 수 없습니다.")
            return

        if not hwp_export.hwp_available():
            self._show_error(
                "한글 필요",
                "한글 자동화(COM)를 사용할 수 없습니다.\n"
                "Windows + 한글(HWP)이 설치되어 있어야 하며, pywin32 패키지가 필요합니다.\n"
                "(pip install pywin32)")
            return

        from utils import resource_path
        template = resource_path(os.path.join("templates", "operation_plan_template.hwp"))
        if not os.path.exists(template):
            template = os.path.join(get_user_data_dir(), "templates", "operation_plan_template.hwp")
        if not os.path.exists(template):
            self._show_error("양식 없음", "운영계획변경 한글 양식 파일을 찾을 수 없습니다.")
            return

        out_path = filedialog.asksaveasfilename(
            title="운영계획변경 저장", defaultextension=".hwp",
            filetypes=[("한글 파일", "*.hwp")],
            initialfile=f"{project_name}_운영계획변경.hwp",
            parent=self,
        )
        if not out_path:
            return

        payload = self._build_hwp_payload(project_name, purpose, unit_price)
        try:
            hwp_export.export_operation_change(template, out_path, payload)
        except Exception as e:
            log_exception("HWP 내보내기 실패")
            self._show_error("내보내기 실패", f"한글 양식 작성 중 오류가 발생했습니다.\n{e}")
            return
        self._show_info("완료", f"한글 양식을 작성했습니다.\n{out_path}")

    # ───────────────────────────── 저장 / 불러오기 ─────────────────────────────
    def _save_business_plan(self, base_dir):
        fpath = os.path.join(base_dir, BUSINESS_PLAN_CSV)
        with open(fpath, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["route", "direction", "lane", "start", "end", "method", "pavement", "year"])
            for it in self.business_plan:
                w.writerow([
                    it.get("route", ""), it.get("direction", ""), it.get("lane", ""),
                    it.get("start", 0), it.get("end", 0), it.get("method", ""),
                    it.get("pavement", ""), it.get("year", self.plan_year),
                ])

    def _load_business_plan(self, base_dir):
        fpath = os.path.join(base_dir, BUSINESS_PLAN_CSV)
        if not os.path.exists(fpath):
            return
        loaded = []

        def _read(enc):
            del loaded[:]
            with open(fpath, "r", encoding=enc) as f:
                for row in csv.DictReader(f):
                    try:
                        loaded.append({
                            "route": (row.get("route") or "").strip(),
                            "direction": (row.get("direction") or "").strip(),
                            "lane": (row.get("lane") or "전차로").strip(),
                            "start": float(row.get("start") or 0),
                            "end": float(row.get("end") or 0),
                            "method": (row.get("method") or "절삭 덧씌우기").strip(),
                            "pavement": (row.get("pavement") or "").strip(),
                            "year": (row.get("year") or self.plan_year).strip(),
                        })
                    except Exception:
                        continue
        try:
            read_csv_with_encoding(_read)
        except Exception:
            log_exception("사업계획 불러오기 실패")
            return
        self.business_plan = loaded
        if self.business_plan:
            self.business_plan_confirmed = True
            yrs = [it.get("year") for it in self.business_plan if it.get("year")]
            if yrs:
                self.plan_year = yrs[0]
            self._inject_plan_entries()

    def _save_operation_changes(self, base_dir):
        fpath = os.path.join(base_dir, OPERATION_CHANGE_CSV)
        with open(fpath, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["route", "direction", "lane", "start", "end", "method",
                        "pavement", "note", "orig_start", "orig_end"])
            for it in self.operation_changes:
                w.writerow([
                    it.get("route", ""), it.get("direction", ""), it.get("lane", ""),
                    it.get("start", 0), it.get("end", 0), it.get("method", ""),
                    it.get("pavement", ""), it.get("note", ""),
                    "" if it.get("orig_start") is None else it.get("orig_start"),
                    "" if it.get("orig_end") is None else it.get("orig_end"),
                ])

    def _load_operation_changes(self, base_dir):
        fpath = os.path.join(base_dir, OPERATION_CHANGE_CSV)
        if not os.path.exists(fpath):
            return
        loaded = []

        def _read(enc):
            del loaded[:]
            with open(fpath, "r", encoding=enc) as f:
                for row in csv.DictReader(f):
                    try:
                        os_v = row.get("orig_start")
                        oe_v = row.get("orig_end")
                        loaded.append({
                            "route": (row.get("route") or "").strip(),
                            "direction": (row.get("direction") or "").strip(),
                            "lane": (row.get("lane") or "전차로").strip(),
                            "start": float(row.get("start") or 0),
                            "end": float(row.get("end") or 0),
                            "method": (row.get("method") or "절삭 덧씌우기").strip(),
                            "pavement": (row.get("pavement") or "").strip(),
                            "note": (row.get("note") or "운영계획 변경").strip(),
                            "orig_start": float(os_v) if os_v not in (None, "") else None,
                            "orig_end": float(oe_v) if oe_v not in (None, "") else None,
                        })
                    except Exception:
                        continue
        try:
            read_csv_with_encoding(_read)
        except Exception:
            log_exception("운영계획변경 불러오기 실패")
            return
        self.operation_changes = loaded
