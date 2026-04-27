# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - 창 종료, 창 크기, 자동 불러오기
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


class WindowMixin:
    def on_closing(self):
        """프로그램 종료 확인 팝업 (버튼 순서: 예-아니오-취소)"""
        # 팝업 윈도우 생성
        dlg = self._create_popup_window(self)
        dlg.title("")
        
        # 윈도우 크기 및 중앙 정렬
        w, h = 360, 150
        x = self.winfo_x() + (self.winfo_width() // 2) - (w // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (h // 2)
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.resizable(False, False)
        
        # 모달 설정 (부모 창 제어 잠금)
        dlg.transient(self)
        dlg.grab_set()
        dlg.focus_force()
        
        # 메시지 라벨
        msg = "변경 사항을 저장하고 종료하시겠습니까?\n(저장하지 않은 데이터는 소실됩니다)"
        ctk.CTkLabel(dlg, text=msg, font=(self.font_family, 13)).pack(pady=20)
        
        # 버튼 프레임
        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        # --- 동작 함수 정의 ---
        def do_save_and_quit():
            dlg.destroy()
            # 파일 대화상자 없이 동일 폴더에 자동 저장 (on_menu_save_route)
            try:
                self.on_menu_save_route()
            except Exception:
                log_exception("종료 전 자동 저장 실패")
            self.destroy()

        def do_quit_no_save():
            dlg.destroy()
            self.destroy()
            
        def do_cancel():
            dlg.destroy()
            
        # --- 버튼 배치 (요청하신 순서: 예 -> 아니오 -> 취소) ---
        
        # 1. 예 (저장 후 종료) - 파란색 강조
        self._create_button(btn_frame, text="예", command=do_save_and_quit, width=70, 
                      fg_color="#3182CE", hover_color="#2B6CB0").pack(side="left", padx=5)
        
        # 2. 아니오 (저장 안 함) - 빨간색 경고
        self._create_button(btn_frame, text="아니오", command=do_quit_no_save, width=70, 
                      fg_color="#E53E3E", hover_color="#C53030").pack(side="left", padx=5)
        
        # 3. 취소 (종료 취소) - 회색 기본
        self._create_button(btn_frame, text="취소", command=do_cancel, width=70, 
                      fg_color="#718096", hover_color="#4A5568").pack(side="left", padx=5)

    def set_initial_window_size(self):
        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            target_w = max(1400, int(sw * 0.85))
            target_h = max(800, int(sh * 0.85))
            x = (sw - target_w) // 2
            y = (sh - target_h) // 2
            self.geometry(f"{target_w}x{target_h}+{x}+{y}")
        except Exception:
            pass

    # ---------- Undo (Ctrl+Z) ----------

    # ---------- 자동 불러오기 ----------
    def auto_load_csvs_on_start(self):
        """실행 시 동일 폴더의 CSV 파일들을 자동으로 불러옵니다."""
        base_dir = get_user_data_dir()

        try:
            if self._load_from_sqlite_cache(base_dir):
                return
        except Exception:
            log_exception("SQLite 캐시 로드 실패")

        # 통합 파일이 존재하면 우선 사용, 없으면 기존 개별 파일 패턴 사용
        def get_files(suffix, all_name):
            all_path = os.path.join(base_dir, all_name)
            if os.path.exists(all_path):
                return [all_name]
            return [fn for fn in os.listdir(base_dir) if fn.endswith(suffix) and not fn.startswith("all_")]

        try:
            files = get_files("_maintenance_history.csv", "all_maintenance_history.csv")
            ic_files = get_files("_ic.csv", "all_ic.csv")
            struct_files = get_files("_structures.csv", "all_structures.csv")
            aar_files = get_files("_aar_data.csv", "all_aar_data.csv")
            branch_files = get_files("_branch_history.csv", "all_branch_history.csv")
            cond_files = get_files("_condition.csv", "all_condition.csv")
            check_cond = "check condition.csv"
            if os.path.exists(os.path.join(base_dir, check_cond)) and check_cond not in cond_files:
                cond_files.append(check_cond)
        except Exception:
            log_exception("자동 로드 대상 파일 수집 실패")
            files, ic_files, struct_files, aar_files, branch_files, cond_files = [], [], [], [], [], []

        if not any([files, ic_files, struct_files, cond_files, branch_files]):
            return  # 자동으로 불러올 파일 없음

        # route_name 기준으로 묶어서 로드
        route_to_rows = {}
        for fname in files:
            fpath = os.path.join(base_dir, fname)
            rname_from_fname = os.path.splitext(fname)[0].replace("_maintenance_history", "")
            def _read_to_group(enc: str):
                with open(fpath, "r", encoding=enc) as f:
                    reader = csv.DictReader(f)
                    read_any = False
                    for row in reader:
                        read_any = True
                        rname = row.get("route_name") or rname_from_fname
                        if rname not in route_to_rows:
                            route_to_rows[rname] = []
                        route_to_rows[rname].append(row)
                    if not read_any:
                        # 헤더만 있는 빈 파일도 노선으로 인식
                        route_to_rows.setdefault(rname_from_fname, [])
            try:
                read_csv_with_encoding(_read_to_group)
            except Exception:
                continue

        # IC 파일들 읽어서 route별로 묶기
        route_to_ics = {}
        for fname in ic_files:
            fpath = os.path.join(base_dir, fname)
            def _read_ic(enc: str):
                with open(fpath, "r", encoding=enc) as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        rname = row.get("route_name") or ""
                        ic_name = row.get("ic_name") or "IC"
                        try:
                            ic_km = float(row.get("ic_km") or 0.0)
                        except Exception:
                            ic_km = 0.0
                        ramp_name = row.get("ramp_name") or ""
                        # ramp_km: 신규 컬럼, 구버전(ramp_start_km) 호환
                        try:
                            rkm = float(row.get("ramp_km") or row.get("ramp_start_km") or ic_km)
                        except Exception:
                            rkm = ic_km
                        route_to_ics.setdefault(rname, []).append({
                            "name": ic_name,
                            "km": ic_km,
                            "ramp": {"name": ramp_name, "km": rkm}
                        })
            try:
                read_csv_with_encoding(_read_ic)
            except Exception:
                continue

        # 지사 이력 파일 읽기: route -> [entries]
        route_to_branch_entries = {}
        for fname in branch_files:
            fpath = os.path.join(base_dir, fname)
            def _read_branch_ent(enc: str):
                with open(fpath, "r", encoding=enc) as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        rname = row.get("route_name") or ""
                        if not rname:
                            continue
                        try:
                            route_to_branch_entries.setdefault(rname, []).append({
                                "start": float(row.get("start_km") or 0.0),
                                "end": float(row.get("end_km") or 0.0),
                                "method": row.get("method", "단면보수"),
                                "direction": (row.get("direction") or "").strip(),
                                "lane": row.get("lane", "전차로"),
                                "work_date": row.get("work_date", ""),
                                "ts": row.get("timestamp", ""),
                            })
                        except Exception:
                            continue
            try:
                read_csv_with_encoding(_read_branch_ent)
            except Exception:
                continue

        # 통합 Condition 데이터 파일 읽기 (HPCI, DI, RD, IRI)
        route_to_di_data = {}
        route_to_hpci_data = {}
        route_to_rd_data = {}
        route_to_iri_data = {}
        route_to_pavement_data = {}

        for fname in cond_files:
            fpath = os.path.join(base_dir, fname)
            # 파일명에서 연도 추출 (예: 2025_condition.csv 또는 2025_평택제천선_condition.csv)
            parts = fname.split('_')
            # '2025년' 등 한글이 포함된 경우도 처리하기 위해 숫자만 추출
            raw_year = parts[0]
            digits = "".join(filter(str.isdigit, raw_year))
            if len(digits) == 4:
                year_str = digits
            else:
                year_str = datetime.now().strftime("%Y")

            # 파일명에서 노선명 추정 (CSV 내 route_name이 비어있을 경우 대비)
            guessed_route_name = ""
            if len(parts) >= 3:
                # 예: 2025_평택제천선_condition.csv -> 평택제천선
                guessed_route_name = "_".join(parts[1:-1])

            def _read_cond(enc: str):
                with open(fpath, "r", newline="", encoding=enc) as f:
                    reader = csv.DictReader(f)

                    # 헤더 공백 제거
                    if reader.fieldnames:
                        reader.fieldnames = [str(name).strip() for name in reader.fieldnames]

                    # 헤더가 올바른지 확인 (노선명 컬럼 존재 여부)
                    expected_keys = ["route_name", "노선명", "route", "노선"]
                    has_header = any(k in (reader.fieldnames or []) for k in expected_keys)

                    rows = []
                    if not has_header:
                        f.seek(0)
                        std_cols = ["route_name", "direction", "lane", "start_km", "end_km", "hpci", "di", "rd", "iri"]
                        reader = csv.DictReader(f, fieldnames=std_cols)
                        rows = list(reader)
                    else:
                        rows = list(reader)

                    for row in rows:
                        clean_row = {k: (v.strip() if v else "") for k, v in row.items() if k}

                        rname = clean_row.get("route_name") or clean_row.get("노선명") or clean_row.get("route") or clean_row.get("노선") or guessed_route_name
                        if not rname:
                            continue

                        d = clean_row.get("direction") or clean_row.get("방향") or ""
                        l = clean_row.get("lane") or clean_row.get("차로") or ""
                        try:
                            skm_val = float(clean_row.get("start_km") or clean_row.get("시작이정") or clean_row.get("start") or 0)
                            skm_str = f"{skm_val:.2f}"
                        except Exception:
                            continue

                        # row의 year 컬럼 우선, 없으면 파일명에서 추출한 year_str 사용 (하위호환)
                        row_year = clean_row.get("year") or clean_row.get("연도") or year_str

                        key = (d, l, skm_str)

                        if clean_row.get("hpci"):
                            try:
                                route_to_hpci_data.setdefault(rname, {}).setdefault(key, {})[row_year] = int(clean_row["hpci"])
                            except ValueError:
                                pass
                        if clean_row.get("di"):
                            try:
                                route_to_di_data.setdefault(rname, {}).setdefault(key, {})[row_year] = float(clean_row["di"])
                            except ValueError:
                                pass
                        if clean_row.get("rd"):
                            try:
                                route_to_rd_data.setdefault(rname, {}).setdefault(key, {})[row_year] = float(clean_row["rd"])
                            except ValueError:
                                pass
                        if clean_row.get("iri"):
                            try:
                                route_to_iri_data.setdefault(rname, {}).setdefault(key, {})[row_year] = float(clean_row["iri"])
                            except ValueError:
                                pass
                        pav = clean_row.get("포장형식", "").strip()
                        if pav:
                            route_to_pavement_data.setdefault(rname, {})[key] = pav
            try:
                read_csv_with_encoding(_read_cond)
            except Exception:
                continue

        # HPCI 데이터 파일 읽기 (제거됨 - 통합 로직으로 대체)
        # ...

        # AAR 데이터 파일 읽기 (유지)
        route_to_aar_data = {}
        for fname in aar_files:
            fpath = os.path.join(base_dir, fname)
            def _read_aar(enc: str):
                with open(fpath, "r", newline="", encoding=enc) as f:
                    reader = csv.DictReader(f)
                    # 헤더 공백 제거
                    if reader.fieldnames:
                        reader.fieldnames = [str(name).strip() for name in reader.fieldnames]

                    for row in reader:
                        # 값 공백 제거
                        clean_row = {k: (v.strip() if v else "") for k, v in row.items() if k}

                        rname = clean_row.get("route_name") or clean_row.get("노선명") or ""
                        if not rname: continue
                        aar_data_for_route = route_to_aar_data.setdefault(rname, {})
                        try:
                            grade_val = clean_row.get('aar_grade') or clean_row.get('AAR등급') or clean_row.get('aar')
                            if grade_val:
                                d = clean_row.get('direction') or clean_row.get('방향') or ""
                                l = clean_row.get('lane') or clean_row.get('차로') or ""
                                s_val = clean_row.get('start_km') or clean_row.get('시작이정') or clean_row.get('start') or "0"
                                key = (d, l, f"{float(s_val):.2f}")
                                year = clean_row.get('year') or clean_row.get('연도') or datetime.now().strftime("%Y")
                                if key not in aar_data_for_route: aar_data_for_route[key] = {}
                                aar_data_for_route[key][year] = int(grade_val)
                        except (KeyError, ValueError):
                            continue
            try:
                read_csv_with_encoding(_read_aar)
            except Exception:
                continue

        # 구조물 데이터 파일 읽기
        route_to_structures = {}
        for fname in struct_files:
            fpath = os.path.join(base_dir, fname)
            def _read_struct(enc: str):
                with open(fpath, "r", newline="", encoding=enc) as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        rname = row.get("route_name") or ""
                        if not rname: continue
                        structs_for_route = route_to_structures.setdefault(rname, [])
                        try:
                            start_km = float(row.get('start_km', 0.0))
                            # length_m이 있으면 사용, 없으면 end_km - start_km (하위 호환)
                            if row.get('length_m'):
                                length_m = float(row['length_m'])
                                if length_m < 0:
                                    # start_km 컬럼에 끝이정이 기재된 경우 (연장이 음수)
                                    end_km = start_km
                                    start_km = end_km + length_m / 1000.0
                                    length_m = abs(length_m)
                                else:
                                    end_km = start_km + length_m / 1000.0
                            else:
                                end_km = float(row.get('end_km', start_km))
                                length_m = (end_km - start_km) * 1000.0
                            
                            direction = (row.get('direction') or '양방향').strip()
                            
                            new_struct = {
                                'name': row.get('name', ''), 
                                'type': row.get('type', ''),
                                'direction': direction,
                                'start': start_km, 
                                'end': end_km,
                                'length': length_m
                            }

                            # 중복 검사 (이름, 시작이정, 방향이 같으면 중복으로 간주)
                            is_duplicate = False
                            for s in structs_for_route:
                                if (s['name'] == new_struct['name'] and 
                                    abs(s['start'] - new_struct['start']) < 1e-6 and
                                    s['direction'] == new_struct['direction']):
                                    is_duplicate = True
                                    break
                            
                            if not is_duplicate:
                                structs_for_route.append(new_struct)
                        except (KeyError, ValueError):
                            continue
            try:
                read_csv_with_encoding(_read_struct)
            except Exception:
                pass # 파일 읽기 실패 시 다음 파일로 넘어감

        # 공법 파일 읽기 (가장 먼저 처리하여 다른 데이터 로드 시 사용 가능하도록)
        load_method_settings(base_dir)

        if not any([route_to_rows, route_to_ics, route_to_branch_entries, route_to_di_data, route_to_hpci_data, route_to_structures]):
            return

        # 이름 → 인덱스 매핑
        name_to_idx = {r["name"]: idx for idx, r in enumerate(self.routes)}

        total_loaded_entries = 0
        created_routes = 0
        
        # 모든 소스에서 노선명을 수집하되, 발견된 순서를 유지 (CSV 작성 순서 반영)
        all_route_names_ordered = []
        seen_routes = set()
        
        def collect_routes_from(source_dict):
            for name in source_dict.keys():
                if name not in seen_routes:
                    seen_routes.add(name)
                    all_route_names_ordered.append(name)

        # 우선순위: 이력 -> IC -> 지사이력 -> 포장상태 -> 구조물
        for d in [route_to_rows, route_to_ics, route_to_branch_entries, route_to_di_data, route_to_hpci_data, route_to_structures, route_to_aar_data, route_to_rd_data, route_to_iri_data]:
            collect_routes_from(d)

        for rname in all_route_names_ordered:
            rows = route_to_rows.get(rname, [])
            # 시작/끝 이정 결정: 히스토리 있으면 첫 행 기준, 없으면 IC 기준(±1km) 또는 기본값
            try:
                if rows:
                    first = rows[0]
                    start_km = float(first.get("route_start_km") or first.get("start_km") or first.get("start") or DEFAULT_START_KM)
                    end_km = float(first.get("route_end_km") or first.get("end_km") or first.get("end") or DEFAULT_END_KM)
                else:
                    ic_list_tmp = route_to_ics.get(rname, [])
                    if ic_list_tmp:
                        kms = [float(rec.get("km") or 0.0) for rec in ic_list_tmp if rec.get("km")]
                        mn, mx = min(kms), max(kms)
                        start_km = max(0.0, math.floor(mn))
                        end_km = math.ceil(mx + 1.0)
                    elif not rows:
                            start_km, end_km = DEFAULT_START_KM, DEFAULT_END_KM
            except Exception:
                start_km, end_km = DEFAULT_START_KM, DEFAULT_END_KM

            entries = []
            for row in rows:
                try:
                    entries.append({
                        "start": float(row.get("start_km") or row.get("start") or 0.0),
                        "end": float(row.get("end_km") or row.get("end") or 0.0),
                        "method": row.get("method", "절삭 덧씌우기"),
                        "direction": (row.get("direction") or "").strip(),
                        "lane": row.get("lane", "전차로"),
                        "work_date": row.get("work_date", ""),
                        "ts": row.get("timestamp", ""),
                    })
                except Exception:
                    continue
            total_loaded_entries += len(entries)

            # 방향 메타 우선 복구 → 없으면 entries에서 유추 → 최후 기본값
            dir1_meta, dir2_meta = None, None
            try:
                for row in rows:
                    d1 = (row.get("route_dir1") or "").strip()
                    d2 = (row.get("route_dir2") or "").strip()
                    if d1 and d2:
                        dir1_meta, dir2_meta = d1, d2
                        break
            except Exception:
                pass
            if dir1_meta and dir2_meta:
                dirs = [dir1_meta, dir2_meta]
            else:
                seen = []
                for it in entries:
                    d = (it.get("direction") or "").strip()
                    if d and d not in seen:
                        seen.append(d)
                    if len(seen) >= 2:
                        break
                if len(seen) == 2:
                    dirs = [seen[0], seen[1]]
                elif len(seen) == 1:
                    other = DIRECTIONS[1] if seen[0] == DIRECTIONS[0] else DIRECTIONS[0]
                    dirs = [seen[0], other]
                else:
                    dirs = DIRECTIONS[:]

            if rname in name_to_idx:
                idx = name_to_idx[rname]
                self.routes[idx]["start_km"] = float(start_km)
                self.routes[idx]["end_km"] = float(end_km)
                self.routes[idx]["entries"] = entries
                self.routes[idx]["directions"] = dirs
            else:
                self.add_route(rname, start_km, end_km, dirs)
                new_idx = self.current_route_index
                self.routes[new_idx]["entries"] = entries
                name_to_idx[rname] = new_idx
                created_routes += 1
            
            # DI 데이터 병합
            if rname in route_to_di_data and rname in name_to_idx:
                self.routes[name_to_idx[rname]]['di_data'] = route_to_di_data[rname]

            # HPCI 데이터 병합 (파일이 없더라도 필드는 존재해야 함)
            if rname in name_to_idx:
                self.routes[name_to_idx[rname]]['hpci_data'] = route_to_hpci_data.get(rname, {})

            # AAR 데이터 병합
            if rname in name_to_idx:
                self.routes[name_to_idx[rname]]['aar_data'] = route_to_aar_data.get(rname, {})

            # RD 데이터 병합
            if rname in name_to_idx:
                self.routes[name_to_idx[rname]]['rd_data'] = route_to_rd_data.get(rname, {})

            # IRI 데이터 병합
            if rname in name_to_idx:
                self.routes[name_to_idx[rname]]['iri_data'] = route_to_iri_data.get(rname, {})

            # 포장형식 데이터 병합
            if rname in name_to_idx:
                self.routes[name_to_idx[rname]]['pavement_data'] = route_to_pavement_data.get(rname, {})

            # 구조물 데이터 병합
            if rname in route_to_structures and rname in name_to_idx:
                self.routes[name_to_idx[rname]]['structures'] = route_to_structures[rname]

            # 지사 이력 병합
            if rname in name_to_idx:
                self.routes[name_to_idx[rname]]['branch_entries'] = route_to_branch_entries.get(rname, [])

            # 이 노선의 IC들도 self.ics에 병합 (히스토리 없어도 처리)
            ic_list = route_to_ics.get(rname, [])
            grouped = {}
            for rec in ic_list:
                key = (rname, rec.get('name'), float(rec.get('km', 0.0)))
                grp = grouped.setdefault(key, {"name": rec.get('name'), "km": float(rec.get('km',0.0)), "route": rname, "ramps": []})
                rp = rec.get('ramp') or {}
                if rp.get('name'):
                    try:
                        rkm = float(rp.get('km') or grp['km'])
                    except Exception:
                        rkm = grp['km']
                    grp["ramps"].append({"name": rp.get('name') or '', "km": rkm})
            # 현재 노선에서 그룹화된 IC/JCT 정보를 self.ics에 누적 추가
            self.ics.extend(list(grouped.values()))

        # 새로 그리기 및 연도 목록 갱신
        for route in self.routes:
            route["_data_version"] = int(route.get("_data_version", 0)) + 1
            route["_year_values_cache"] = None
            route["_year_values_version"] = -1
        self._mark_route_data_changed()
        self.update_year_filter_values()
        self.draw_schematic()
        
        # 공법 콤보박스 업데이트 및 범례 다시 그리기 (자동 로드 후 UI 강제 갱신)
        try:
            self.update_method_combobox_values()
            self.render_legend()
        except Exception:
            log_exception("자동 로드 후 UI 갱신 실패")

    # ---------- 그리기 ----------
