# -*- coding: utf-8 -*-
"""모식도에서 직접 구간을 선택하는 '모식도 선택 모드' 믹스인.

사업대상 선정(사업계획)·운영계획변경에서 '개량 우선순위 산정' 외에
모식도를 직접 보고 구간을 고를 수 있게 한다.

동작 개요
─────────
1) plan_mixin 의 버튼('모식도에서 선택')이 ``begin_schematic_selection`` 호출
2) 메인 모식도가 자동으로 DI 지수(포장상태불량률) 색상 보기로 전환
3) 보수이력이 있는 구간 + 교량/터널 구간은 어둡게(빗금) 표시되어 선택 불가
4) 상행/하행 바의 '차로별'로 드래그하여 km 구간을 선택(100m 스냅)
   - 드래그 안에 선택불가 구간이 섞이면 자동으로 잘라냄(나머지만 선택)
   - Shift 를 누른 채 드래그하면 기존 선택에 '추가'(없으면 새로 대체)
   - 선택된 구간을 그냥 클릭하면 해당 선택을 해제(토글)
5) 떠있는 패널에서 '선택 적용' → 우선순위 산정과 동일한 apply_callback 으로 전달

선택 결과는 ``self.schematic_selections`` (노선명 → [{direction,lane,start,end}]) 에
저장되며, 적용 시 우선순위 팝업의 항목과 동일한 형식(dict)으로 변환해 콜백에 넘긴다.
"""
import tkinter as tk
import customtkinter as ctk

from constants import GRID_KM, PRIMARY_BLUE, PRIMARY_BLUE_HOVER, TITLE_TEXT
from utils import (
    km_floor_to_grid, km_ceil_to_grid, lane_conflict, log_exception,
)
from canvas_utils import _split_gaps, _draw_hatch_rect


class SelectMixin:
    # ───────────────────────────── 진입 / 종료 ─────────────────────────────
    def begin_schematic_selection(self, label, apply_callback,
                                  exclude_sections=None, origin_window=None):
        """모식도 선택 모드 시작.

        label          : 패널 제목용 문자열(예: '사업대상' / '운영계획변경')
        apply_callback : 적용 시 호출. callback(items:list, plan_year:str)
                         items 는 우선순위 항목과 동일 형식의 dict 리스트
        exclude_sections: 선택 불가로 추가 처리할 구간(운영계획변경의 사업계획 구간 등)
        origin_window  : 선택 동안 잠시 숨겼다가 종료 시 되살릴 창
        """
        if self.current_route_index < 0 or not self.routes:
            self._show_info("알림", "먼저 노선을 추가해 주세요.")
            return
        # 돋보기 모드와 충돌 방지
        if getattr(self, "magnifier_mode", False):
            try:
                self.exit_magnifier_mode()
            except Exception:
                pass

        self.schematic_select_mode = True
        self.schematic_select_label = label
        self.schematic_select_callback = apply_callback
        self.schematic_select_exclude = exclude_sections or []
        self.schematic_select_origin_window = origin_window
        self.schematic_selections = {}
        self._sel_drag = None

        # 보기 상태 백업 후 DI 보기로 전환
        self._sel_view_backup = {
            "mode": self.view_mode.get(),
            "hpci": self.view_hpci.get(), "di": self.view_di.get(),
            "aar": self.view_aar.get(), "rd": self.view_rd.get(),
            "iri": self.view_iri.get(),
        }
        self.view_hpci.set(False)
        self.view_di.set(True)
        self.view_aar.set(False)
        self.view_rd.set(False)
        self.view_iri.set(False)

        if origin_window is not None:
            try:
                origin_window.withdraw()
            except Exception:
                pass

        self._bind_selection_events()
        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass
        self.draw_schematic()
        self._open_selection_panel()

    def _end_schematic_selection(self, restore_origin=True):
        """선택 모드 종료(보기 복원, 패널/바인딩 정리)."""
        self.schematic_select_mode = False
        self._sel_drag = None
        self._unbind_selection_events()

        b = getattr(self, "_sel_view_backup", None)
        if b:
            self.view_hpci.set(b["hpci"])
            self.view_di.set(b["di"])
            self.view_aar.set(b["aar"])
            self.view_rd.set(b["rd"])
            self.view_iri.set(b["iri"])
            self.view_mode.set(b["mode"])
        self._sel_view_backup = None

        panel = getattr(self, "_sel_panel", None)
        if panel is not None:
            try:
                panel.destroy()
            except Exception:
                pass
            self._sel_panel = None
            self._sel_count_lbl = None

        for r in self.routes:
            c = r.get("canvas")
            if c is not None:
                try:
                    c.delete("sel_live")
                except Exception:
                    pass
        self.draw_schematic()

        if restore_origin and self.schematic_select_origin_window is not None:
            try:
                self.schematic_select_origin_window.deiconify()
                self.schematic_select_origin_window.lift()
                self.schematic_select_origin_window.focus_force()
            except Exception:
                pass

    def _cancel_schematic_selection(self):
        self.schematic_selections = {}
        self._end_schematic_selection(restore_origin=True)

    def _apply_schematic_selection(self):
        items = []
        for rname, sels in self.schematic_selections.items():
            route = next((r for r in self.routes if r.get("name") == rname), None)
            for s in sels:
                pav = ""
                if route is not None:
                    pav = self._majority_pavement(route, s["direction"], s["start"], s["end"])
                items.append({
                    "route": rname,
                    "direction": s["direction"],
                    "lane": s["lane"],
                    "start": round(float(s["start"]), 3),
                    "end": round(float(s["end"]), 3),
                    "method": "절삭 덧씌우기",
                    "pavement": pav,
                })
        if not items:
            self._show_info("알림", "선택된 구간이 없습니다.\n모식도에서 구간을 드래그해 선택해 주세요.")
            return
        cb = self.schematic_select_callback
        self.schematic_selections = {}
        self._end_schematic_selection(restore_origin=True)
        if callable(cb):
            try:
                cb(items, self.plan_year)
            except Exception:
                log_exception("모식도 선택 적용 콜백 실패")

    # ───────────────────────────── 이벤트 바인딩 ─────────────────────────────
    def _bind_selection_events(self):
        for r in self.routes:
            c = r.get("canvas")
            if c is None:
                continue
            c.bind("<Button-1>", self._on_selection_press)
            c.bind("<B1-Motion>", self._on_selection_motion)
            c.bind("<ButtonRelease-1>", self._on_selection_release)

    def _unbind_selection_events(self):
        for r in self.routes:
            c = r.get("canvas")
            if c is None:
                continue
            try:
                c.bind("<Button-1>", lambda e: e.widget.focus_set())
                c.unbind("<B1-Motion>")
                c.unbind("<ButtonRelease-1>")
                c.delete("sel_live")
            except Exception:
                pass

    # ───────────────────────────── 마우스 인터랙션 ─────────────────────────────
    def _on_selection_press(self, event):
        if not getattr(self, "schematic_select_mode", False):
            return
        canvas = event.widget
        try:
            canvas.focus_set()
        except Exception:
            pass
        if canvas not in self.canvas_to_route_index:
            return
        ridx = self.canvas_to_route_index[canvas]
        self.current_route_index = ridx
        route = self.routes[ridx]
        cx = canvas.canvasx(event.x)
        cy = canvas.canvasy(event.y)
        dl = self._dir_lane_at(route, cy)
        mapping = route.get("_km_mapping")
        if dl is None or not mapping:
            self._sel_drag = None
            return
        direction, lane_str, y1, y2 = dl
        self._sel_drag = {
            "route_idx": ridx, "direction": direction, "lane": lane_str,
            "start_km": self._px_to_km(cx, mapping), "start_cx": cx,
            "y1": y1, "y2": y2,
        }

    def _on_selection_motion(self, event):
        drag = getattr(self, "_sel_drag", None)
        if not drag:
            return
        canvas = event.widget
        canvas.delete("sel_live")
        route = self.routes[drag["route_idx"]]
        mapping = route.get("_km_mapping")
        if not mapping:
            return
        cur_km = self._px_to_km(canvas.canvasx(event.x), mapping)
        lo = km_floor_to_grid(min(drag["start_km"], cur_km), GRID_KM)
        hi = km_ceil_to_grid(max(drag["start_km"], cur_km), GRID_KM)
        r_start = float(route["start_km"]); r_end = float(route["end_km"])
        lo = max(lo, r_start); hi = min(hi, r_end)
        if hi <= lo:
            return
        self._draw_band_px(
            canvas, route, lo, hi, drag["y1"] + 1, drag["y2"] - 1,
            dict(fill="#3B82F6", stipple="gray25", outline="#1D4ED8",
                 width=2, tags=("sel_live",)))
        canvas.tag_raise("sel_live")

    def _on_selection_release(self, event):
        drag = getattr(self, "_sel_drag", None)
        self._sel_drag = None
        if not drag:
            return
        canvas = event.widget
        canvas.delete("sel_live")
        route = self.routes[drag["route_idx"]]
        mapping = route.get("_km_mapping")
        if not mapping:
            return
        cx = canvas.canvasx(event.x)
        end_km = self._px_to_km(cx, mapping)
        direction = drag["direction"]; lane_str = drag["lane"]
        shift = bool(event.state & 0x0001)

        # 단순 클릭(이동량 작음) → 기존 선택 해제(토글)
        if abs(cx - drag["start_cx"]) < 5:
            self._remove_selection_at(route, direction, lane_str, end_km)
            self.draw_schematic()
            self._update_selection_panel()
            return

        lo = km_floor_to_grid(min(drag["start_km"], end_km), GRID_KM)
        hi = km_ceil_to_grid(max(drag["start_km"], end_km), GRID_KM)
        r_start = float(route["start_km"]); r_end = float(route["end_km"])
        lo = max(lo, r_start); hi = min(hi, r_end)
        if hi - lo <= 1e-9:
            return

        # 선택불가 구간 자동 잘라내기
        blocked = self._selection_blocked_intervals(route, direction, lane_str)
        free = [(s, e) for s, e in self._subtract_intervals(lo, hi, blocked)
                if e - s > 1e-6]

        # 드래그 전체가 선택불가면 기존 선택을 지우지 않고 무시
        if not free:
            return
        if not shift:
            self.schematic_selections = {}
        lst = self.schematic_selections.setdefault(route["name"], [])
        for s, e in free:
            lst.append({"direction": direction, "lane": lane_str,
                        "start": round(s, 3), "end": round(e, 3)})
        self._normalize_selections(route["name"])
        self.draw_schematic()
        self._update_selection_panel()

    # ───────────────────────────── 좌표 / 차로 판정 ─────────────────────────────
    def _dir_lane_at(self, route, cy):
        """캔버스 y좌표 → (방향, '차로' 문자열, 차로 y1, y2). 바 밖이면 None."""
        bl = route.get("_bar_layout")
        if not bl:
            return None
        bar_h = bl["bar_h"]
        lane_count = route.get("lane_count", 4)
        if lane_count <= 0:
            return None
        seg_h = bar_h / lane_count
        directions = route.get("directions", [])
        for di in range(min(2, len(directions))):
            top = bl["bar1_top"] if di == 0 else bl["bar2_top"]
            if top <= cy <= top + bar_h:
                seg_from_top = int((cy - top) // seg_h)
                seg_from_top = max(0, min(lane_count - 1, seg_from_top))
                # 1차로는 중분대(중앙)에 인접: 상행 바는 아래쪽, 하행 바는 위쪽이 1차로
                n = (lane_count - seg_from_top) if di == 0 else (seg_from_top + 1)
                if di == 0:
                    y1 = top + bar_h - n * seg_h
                    y2 = top + bar_h - (n - 1) * seg_h
                else:
                    y1 = top + (n - 1) * seg_h
                    y2 = top + n * seg_h
                return (directions[di], f"{n}차로", y1, y2)
        return None

    def _lane_y_band(self, route, direction, lane_str):
        """(방향, '차로' 문자열) → (y1, y2). 실패 시 None."""
        bl = route.get("_bar_layout")
        if not bl:
            return None
        bar_h = bl["bar_h"]
        lane_count = route.get("lane_count", 4)
        if lane_count <= 0:
            return None
        seg_h = bar_h / lane_count
        directions = route.get("directions", [])
        di = 0 if (directions and direction == directions[0]) else 1
        top = bl["bar1_top"] if di == 0 else bl["bar2_top"]
        try:
            n = int(str(lane_str).replace("차로", ""))
        except (ValueError, AttributeError):
            n = 1
        n = max(1, min(n, lane_count))
        if di == 0:
            return (top + bar_h - n * seg_h, top + bar_h - (n - 1) * seg_h)
        return (top + (n - 1) * seg_h, top + n * seg_h)

    # ───────────────────────────── 선택불가 구간 산정 ─────────────────────────────
    def _selection_blocked_intervals(self, route, direction, lane_str):
        """해당 (방향,차로)에서 선택할 수 없는 km 구간 목록(병합·정렬)."""
        blocked = []

        # 교량/터널 등 구조물: 방향이 맞으면 모든 차로 차단
        for s in route.get("structures", []):
            s_dir = (s.get("direction") or "양방향").strip()
            if s_dir != "양방향" and (s_dir not in direction) and (direction not in s_dir):
                continue
            try:
                a, b = float(s["start"]), float(s["end"])
            except (ValueError, TypeError, KeyError):
                continue
            if abs(b - a) < 1e-6:
                b = a + 0.1
            if b > a:
                blocked.append((a, b))

        # 보수이력(계획 포함): 방향·차로가 충돌하면 차단
        for e in route.get("entries", []):
            e_dir = (e.get("direction") or "").strip()
            if e_dir and e_dir != "양방향" and (e_dir not in direction) and (direction not in e_dir):
                continue
            if not lane_conflict(lane_str, e.get("lane", "전차로")):
                continue
            try:
                a, b = float(e["start"]), float(e["end"])
            except (ValueError, TypeError, KeyError):
                continue
            if b > a:
                blocked.append((a, b))

        # 추가 제외 구간(운영계획변경의 사업계획 지정 구간 등)
        for ex in getattr(self, "schematic_select_exclude", []) or []:
            if str(ex.get("route")) != str(route.get("name")):
                continue
            ex_dir = str(ex.get("direction") or "").strip()
            if ex_dir and ex_dir != "양방향" and (ex_dir not in direction) and (direction not in ex_dir):
                continue
            if not lane_conflict(lane_str, ex.get("lane", "전차로")):
                continue
            try:
                a, b = float(ex.get("start", 0)), float(ex.get("end", 0))
            except (ValueError, TypeError):
                continue
            if b > a:
                blocked.append((a, b))

        return self._merge_km_intervals(blocked)

    @staticmethod
    def _merge_km_intervals(intervals):
        """겹치거나 인접한 km 구간 병합."""
        ivs = sorted((min(a, b), max(a, b)) for a, b in intervals)
        out = []
        for a, b in ivs:
            if out and a <= out[-1][1] + 1e-9:
                out[-1] = (out[-1][0], max(out[-1][1], b))
            else:
                out.append([a, b])
        return [(a, b) for a, b in out]

    @staticmethod
    def _subtract_intervals(lo, hi, blocked):
        """[lo,hi] 에서 blocked(병합·정렬) 구간을 뺀 자유 구간 목록."""
        free = []
        cur = lo
        for a, b in blocked:
            if b <= cur:
                continue
            if a >= hi:
                break
            if a > cur:
                free.append((cur, min(a, hi)))
            cur = max(cur, b)
        if cur < hi:
            free.append((cur, hi))
        return free

    def _majority_pavement(self, route, direction, s, e):
        """구간 내 최빈 포장형식(없으면 '')."""
        pav = route.get("pavement_data", {}) or {}
        counter = {}
        for key, v in pav.items():
            try:
                d, _l, skm = key
            except Exception:
                continue
            if d and d != direction and d not in direction and direction not in d:
                continue
            try:
                km = float(skm)
            except (ValueError, TypeError):
                continue
            if s - 1e-9 <= km < e + 1e-9 and v:
                counter[v] = counter.get(v, 0) + 1
        return max(counter, key=counter.get) if counter else ""

    # ───────────────────────────── 선택 목록 관리 ─────────────────────────────
    def _normalize_selections(self, rname):
        """동일 (방향,차로) 내 겹치는 선택을 병합."""
        sels = self.schematic_selections.get(rname)
        if not sels:
            return
        groups = {}
        for s in sels:
            groups.setdefault((s["direction"], s["lane"]), []).append((s["start"], s["end"]))
        new = []
        for (d, l), ivs in groups.items():
            for a, b in self._merge_km_intervals(ivs):
                new.append({"direction": d, "lane": l,
                            "start": round(a, 3), "end": round(b, 3)})
        self.schematic_selections[rname] = new

    def _remove_selection_at(self, route, direction, lane_str, km):
        rname = route["name"]
        sels = self.schematic_selections.get(rname)
        if not sels:
            return
        self.schematic_selections[rname] = [
            s for s in sels
            if not (s["direction"] == direction and s["lane"] == lane_str
                    and s["start"] - 1e-9 <= km <= s["end"] + 1e-9)
        ]

    # ───────────────────────────── 그리기 ─────────────────────────────
    def _draw_band_px(self, canvas, route, s_km, e_km, y1, y2, kw,
                      hatch=False, hatch_color="#9CA3AF"):
        """km 구간을 IC 갭으로 분절해 사각형(+선택 시 빗금)으로 그린다."""
        if e_km <= s_km:
            return
        x1 = self._km_to_canvas_x(route, s_km)
        x2 = self._km_to_canvas_x(route, e_km)
        if x2 < x1:
            x1, x2 = x2, x1
        gaps = route.get("_gaps_px", [])
        segs = _split_gaps(gaps, x1, x2) if gaps else [(x1, x2)]
        for sx1, sx2 in segs:
            if sx2 <= sx1:
                continue
            if kw:
                canvas.create_rectangle(sx1, y1, sx2, y2, **kw)
            if hatch:
                _draw_hatch_rect(canvas, sx1, y1, sx2, y2, color=hatch_color)

    def _draw_schematic_selection_overlay(self, route):
        """선택 모드일 때 선택불가(어둡게)·선택됨(초록) 오버레이를 그린다."""
        if not getattr(self, "schematic_select_mode", False):
            return
        canvas = route.get("canvas")
        bl = route.get("_bar_layout")
        if not canvas or not bl:
            return
        bar_h = bl["bar_h"]
        lane_count = route.get("lane_count", 4)
        if lane_count <= 0:
            return
        seg_h = bar_h / lane_count
        directions = route.get("directions", [])

        # 선택불가 구간(차로별)
        for di in range(min(2, len(directions))):
            direction = directions[di]
            top = bl["bar1_top"] if di == 0 else bl["bar2_top"]
            for n in range(1, lane_count + 1):
                if di == 0:
                    y1 = top + bar_h - n * seg_h
                    y2 = top + bar_h - (n - 1) * seg_h
                else:
                    y1 = top + (n - 1) * seg_h
                    y2 = top + n * seg_h
                lane_str = f"{n}차로"
                for a, b in self._selection_blocked_intervals(route, direction, lane_str):
                    self._draw_band_px(
                        canvas, route, a, b, y1 + 1, y2 - 1,
                        dict(fill="#374151", stipple="gray50", outline="",
                             tags=("sel_blocked",)),
                        hatch=True, hatch_color="#D1D5DB")

        # 선택된 구간(초록)
        for sel in self.schematic_selections.get(route["name"], []):
            band = self._lane_y_band(route, sel["direction"], sel["lane"])
            if band is None:
                continue
            y1, y2 = band
            self._draw_band_px(
                canvas, route, sel["start"], sel["end"], y1 + 1, y2 - 1,
                dict(fill="#22C55E", stipple="gray50", outline="#15803D",
                     width=2, tags=("sel_chosen",)))

        canvas.tag_raise("sel_blocked")
        canvas.tag_raise("sel_chosen")

    # ───────────────────────────── 컨트롤 패널 ─────────────────────────────
    def _open_selection_panel(self):
        panel = self._create_popup_window(self)
        self._sel_panel = panel
        panel.title("")
        try:
            panel.attributes("-topmost", True)
        except Exception:
            pass
        panel.resizable(False, False)
        panel.geometry("480x340")

        label = getattr(self, "schematic_select_label", "")
        ctk.CTkLabel(panel, text=f"모식도에서 {label} 구간 선택",
                     font=(self.font_family, 16, "bold"),
                     text_color=TITLE_TEXT).pack(anchor="w", padx=18, pady=(16, 4))

        ctk.CTkLabel(
            panel, justify="left", anchor="w", wraplength=440,
            font=(self.font_family, 12), text_color="#3A4A60",
            text=(
                "• 색상은 DI 지수(포장상태불량률)입니다. 진할수록 불량합니다.\n"
                "• 상행/하행 바의 차로를 가로로 드래그하면 그 차로 구간이 선택됩니다.\n"
                "• 드래그 안에 선택불가 구간이 있으면 자동으로 잘라냅니다.\n"
                "• Shift 를 누른 채 드래그하면 기존 선택에 추가됩니다.\n"
                "• 선택한 구간을 클릭하면 선택이 해제됩니다."
            ),
        ).pack(anchor="w", padx=18, pady=(0, 8))

        legend = ctk.CTkFrame(panel, fg_color="transparent")
        legend.pack(fill="x", padx=18, pady=(0, 4))
        ctk.CTkLabel(legend, text="■", text_color="#22C55E",
                     font=(self.font_family, 13, "bold")).pack(side="left")
        ctk.CTkLabel(legend, text="선택됨", font=(self.font_family, 11),
                     text_color="#3A4A60").pack(side="left", padx=(2, 14))
        ctk.CTkLabel(legend, text="▦", text_color="#6B7280",
                     font=(self.font_family, 13, "bold")).pack(side="left")
        ctk.CTkLabel(legend, text="선택불가(보수이력·교량·터널)",
                     font=(self.font_family, 11),
                     text_color="#3A4A60").pack(side="left", padx=(2, 0))

        self._sel_count_lbl = ctk.CTkLabel(
            panel, text="", font=(self.font_family, 13, "bold"),
            text_color=PRIMARY_BLUE)
        self._sel_count_lbl.pack(anchor="w", padx=18, pady=(8, 4))

        btns = ctk.CTkFrame(panel, fg_color="transparent")
        btns.pack(fill="x", padx=18, pady=(6, 14))
        self._create_button(
            btns, text="선택 적용", command=self._apply_schematic_selection,
            width=120, fg_color="#2F855A", hover_color="#276749",
            text_color="#FFFFFF",
            font=(self.font_family, 13, "bold")).pack(side="right", padx=4)
        self._create_button(
            btns, text="전체 해제", command=self._clear_schematic_selection,
            width=100, fg_color="#A0AEC0", hover_color="#8A98AB").pack(side="left", padx=2)
        self._create_close_button(
            btns, self._cancel_schematic_selection, width=90,
            text="취소").pack(side="right", padx=4)

        panel.protocol("WM_DELETE_WINDOW", self._cancel_schematic_selection)
        panel.bind("<Escape>", lambda e: self._cancel_schematic_selection())

        try:
            # 모식도(중앙 바)를 가리지 않도록 우측 상단에 배치
            self.update_idletasks()
            px = self.winfo_rootx() + max(20, self.winfo_width() - 500)
            py = self.winfo_rooty() + 60
            panel.geometry(f"480x340+{px}+{py}")
        except Exception:
            pass
        self._update_selection_panel()

    def _clear_schematic_selection(self):
        self.schematic_selections = {}
        self.draw_schematic()
        self._update_selection_panel()

    def _update_selection_panel(self):
        lbl = getattr(self, "_sel_count_lbl", None)
        if lbl is None:
            return
        cnt = 0
        total = 0.0
        for sels in self.schematic_selections.values():
            for s in sels:
                cnt += 1
                total += max(0.0, float(s["end"]) - float(s["start"]))
        try:
            lbl.configure(text=f"선택된 구간: {cnt}개  ·  총 연장 {total:.2f} km")
        except Exception:
            pass
