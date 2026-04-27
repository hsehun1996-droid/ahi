# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - 유틸리티 함수
"""
import os
import sys
import csv
import math
import json
import logging
from datetime import datetime

from constants import (
    GRID_KM, ENTRY_GRID_KM, METHOD_STYLES, DEFAULT_METHOD_STYLES,
    METHOD_CATEGORY_MAP, CATEGORY_WARRANTY, DEFAULT_LOG_FILE,
)


def resource_path(relative_path):
    """ PyInstaller로 패키징된 리소스 경로를 반환 """
    try:
        # PyInstaller가 생성한 임시 폴더 경로
        base_path = sys._MEIPASS
    except Exception:
        # 일반 Python 스크립트 실행 시 경로
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_user_data_dir():
    """사용자 데이터(CSV 등) 저장/불러오기 기준 폴더를 반환.
    패키징 시: exe가 있는 폴더, 스크립트 실행 시: 스크립트 폴더.
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def read_csv_with_encoding(read_func):
    """utf-8-sig로 먼저 시도하고 실패하면 cp949로 재시도합니다."""
    try:
        read_func("utf-8-sig")
    except UnicodeDecodeError:
        read_func("cp949")


def get_logger():
    logger = logging.getLogger("highway")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    logger.propagate = False
    try:
        log_path = os.path.join(get_user_data_dir(), DEFAULT_LOG_FILE)
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(handler)
    except Exception:
        logger.addHandler(logging.NullHandler())
    return logger


def log_exception(message: str):
    get_logger().exception(message)


def log_warning(message: str):
    get_logger().warning(message)


def reset_method_settings():
    METHOD_STYLES.clear()
    METHOD_STYLES.update({name: style.copy() for name, style in DEFAULT_METHOD_STYLES.items()})
    METHOD_CATEGORY_MAP.clear()
    CATEGORY_WARRANTY.clear()


def get_method_warranty_period(method: str) -> int:
    category = METHOD_CATEGORY_MAP.get(method, "")
    if category and category in CATEGORY_WARRANTY:
        try:
            return int(CATEGORY_WARRANTY[category]["period"])
        except (TypeError, ValueError, KeyError):
            log_warning(f"하자기간 설정이 잘못되어 기본값을 사용합니다: category={category!r}")
    return 2 if category == "표면개량" else 3


def load_method_settings(base_dir: str):
    loaded_methods = {}
    loaded_categories = {}
    loaded_warranty = {}

    method_file_path = os.path.join(base_dir, "all_routes_methods.csv")
    if os.path.exists(method_file_path):
        def _read_methods(enc: str):
            with open(method_file_path, "r", encoding=enc) as fm:
                reader = csv.DictReader(fm)
                for row in reader:
                    name = (row.get("method_name") or "").strip()
                    color = (row.get("fill_color") or "#718096").strip()
                    category = (row.get("category") or "").strip()
                    if not name:
                        continue
                    loaded_methods[name] = {"fill": color}
                    if category:
                        loaded_categories[name] = category
        try:
            read_csv_with_encoding(_read_methods)
        except Exception:
            log_exception(f"공법 설정 로드 실패: {method_file_path}")

    warranty_file_path = os.path.join(base_dir, "warranty_settings.csv")
    if os.path.exists(warranty_file_path):
        def _read_warranty(enc: str):
            with open(warranty_file_path, "r", encoding=enc) as fw:
                reader = csv.DictReader(fw)
                for row in reader:
                    category = (row.get("category") or "").strip()
                    if not category:
                        continue
                    try:
                        period = int(row.get("period", 3))
                    except ValueError:
                        period = 3
                    try:
                        rate = float(row.get("rate", 100.0))
                    except ValueError:
                        rate = 100.0
                    loaded_warranty[category] = {"period": period, "rate": rate}
        try:
            read_csv_with_encoding(_read_warranty)
        except Exception:
            log_exception(f"하자기간 설정 로드 실패: {warranty_file_path}")

    reset_method_settings()
    if loaded_methods:
        METHOD_STYLES.clear()
        METHOD_STYLES.update(loaded_methods)
    METHOD_CATEGORY_MAP.update(loaded_categories)
    CATEGORY_WARRANTY.update(loaded_warranty)


def _safe_stat_signature(path: str):
    try:
        stat = os.stat(path)
        return {
            "name": os.path.basename(path),
            "size": int(stat.st_size),
            "mtime_ns": int(stat.st_mtime_ns),
        }
    except OSError:
        return None


def build_source_signature(base_dir: str) -> str:
    try:
        names = sorted(
            fn for fn in os.listdir(base_dir)
            if fn.lower().endswith(".csv")
        )
    except OSError:
        names = []
    rows = []
    for name in names:
        sig = _safe_stat_signature(os.path.join(base_dir, name))
        if sig is not None:
            rows.append(sig)
    return json.dumps(rows, ensure_ascii=False, sort_keys=True)


def km_floor_to_grid(x: float, grid: float = GRID_KM) -> float:
    return math.floor((x + 1e-9) / grid) * grid


def km_ceil_to_grid(x: float, grid: float = GRID_KM) -> float:
    return math.ceil((x - 1e-9) / grid) * grid


def clamp(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


def fmt_km(x: float) -> str:
    return f"{x:.2f}"


def intervals_overlap(a1, a2, b1, b2) -> bool:
    """[a1, a2)와 [b1, b2) 구간 겹침 여부"""
    return not (a2 <= b1 or b2 <= a1)


def lane_conflict(new_lane: str, exist_lane: str) -> bool:
    """전차로는 모든 차로와 충돌로 간주, 동일 차로는 충돌"""
    if new_lane == "전차로" or exist_lane == "전차로":
        return True
    return new_lane == exist_lane
