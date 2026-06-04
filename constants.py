# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - 상수 및 설정
"""
# -------------------- 설정 --------------------
DEFAULT_ROUTE_NAME = "샘플"
DEFAULT_START_KM = 35.89
DEFAULT_END_KM = 126.91
GRID_KM = 0.1  # 100 m (표시용 그리드)
# 이력 입력/편집 스냅 격자(1 m)
ENTRY_GRID_KM = 0.001

# 모식도 가로 스케일(px/km). 값이 클수록 길게 보입니다.
# 100 m(0.1 km) 폭을 기존 대비 2배로 확대 → 0.1 km 당 16 px
PX_PER_KM = 160

# 공법 색상 정의
DEFAULT_METHOD_STYLES = {
    "절삭 덧씌우기": {"fill": "#2B6CB0"},  # 파랑
}
METHOD_STYLES = {name: style.copy() for name, style in DEFAULT_METHOD_STYLES.items()}

BRANCH_METHOD_STYLES = {
    "단면보수": {"fill": "#808080"},  # 회색
    "패칭":    {"fill": "#000000"},  # 검정
}

# 공법 카테고리 맵: {공법명: 카테고리명}  (카테고리 없으면 키 없음)
METHOD_CATEGORY_MAP: dict = {}

# 카테고리별 하자기간 설정: {카테고리명: {"period": int, "rate": float}}
CATEGORY_WARRANTY: dict = {}
CACHE_SCHEMA_VERSION = 2
DEFAULT_LOG_FILE = "highway.log"

BACKGROUND_FILL = "#F9FAFB"  # 아주 연한 회색 배경
AXIS_COLOR = "#A0AEC0"
TEXT_COLOR = "#2D3748"       # 진한 회색/검정 텍스트
GRID_100M = "#FF0000"        # 100m 점선 (빨강)
GRID_1KM_COLOR = "#000000"   # 1km 눈금선 (검정)
LANE_BORDER = "#000000"      # 차로 구분선 (검정)
IC_COLOR = "#2B6CB0"         # IC 마커 색상 (진한 파랑)
IC_BOX_W = 40             # IC 박스 가로폭(px) - 세로쓰기를 위해 축소
IC_TICK_SPAN_KM = 0.5     # IC 주변 강조 눈금 표시 범위(±km)
IC_GAP_MARGIN = 14        # IC 박스 좌우 여백(px) - 모식도 분리 효과
RAMP_BAR_W = 520          # 램프 미니 모식도 가로폭(px)
RAMP_BAR_H = 22           # 램프 레인 높이(px) - 기존 대비 3/5로 축소(≈21.6→22)
RAMP_BAR_GAP_Y = 36       # 램프 간 수직 간격(px) - 2배 확대
RAMP_PX_PER_KM = 100      # 램프 모식도에서 1.0km당 픽셀 수 (100m 폭 절반)
RAMP_LANE_GAP_Y = 6       # 램프 내부 차로 바 사이 수직 간격(px)

DIRECTIONS = ["상행", "하행"]
LANES = ["전차로", "1차로", "2차로", "3차로", "4차로"]  # 필요 시 갯수 늘려 사용

# -------------------- 캔버스 레이아웃 (scale 1.0 기준) --------------------
CANVAS_WIDGET_H       = 360   # 모식도 캔버스 위젯 높이 (px)
DETAIL_CANVAS_H       = 205   # 돋보기 상세 캔버스 위젯 높이 (px)
CANVAS_BAR_H          = 110   # 방향 바 높이 (px)
CANVAS_TOP_MARGIN     = 60    # 모식도 상단 여백 (px)
CANVAS_DIR_GAP        = 20    # 상행/하행 바 사이 간격 (px)
CANVAS_KM_LBL_OFFSET  = 16    # km 눈금 라벨의 바 외곽 여백 (px)
CANVAS_DIR_LBL_OFFSET = 36    # 방향 라벨의 바 외곽 여백 (px)
CANVAS_IC_CHAR_H      = 16    # IC 세로 텍스트 문자 높이 (px)

# 캔버스 폰트 크기 (캔버스 tk.Canvas용)
CANVAS_FONT_XS  = 7
CANVAS_FONT_S   = 8
CANVAS_FONT_M   = 9
CANVAS_FONT_L   = 10
CANVAS_FONT_XL  = 11

# 현재 적용된 디스플레이 스케일 (apply_display_scale 호출 시 갱신)
DISPLAY_SCALE = 1.0
# 스케일 기준 해상도 너비. 이 값에서 스케일=1.0, 그 외엔 (해상도/REFERENCE_WIDTH) 비율 적용.
REFERENCE_WIDTH = 1920

# ---- 스케일 기준값 (apply_display_scale 에서 참조) ----
_BASE_PX_PER_KM          = 160
_BASE_CANVAS_WIDGET_H    = 360
_BASE_DETAIL_CANVAS_H    = 205
_BASE_RAMP_BAR_W         = 520
_BASE_RAMP_BAR_H         = 22
_BASE_RAMP_BAR_GAP_Y     = 36
_BASE_RAMP_PX_PER_KM     = 100
_BASE_RAMP_LANE_GAP_Y    = 6
_BASE_IC_BOX_W           = 40
_BASE_IC_GAP_MARGIN      = 14
_BASE_CANVAS_BAR_H       = 110
_BASE_CANVAS_TOP_MARGIN  = 60
_BASE_CANVAS_DIR_GAP     = 20
_BASE_CANVAS_KM_LBL_OFF  = 16
_BASE_CANVAS_DIR_LBL_OFF = 36
_BASE_CANVAS_IC_CHAR_H   = 16
_BASE_CANVAS_FONT_XS     = 7
_BASE_CANVAS_FONT_S      = 8
_BASE_CANVAS_FONT_M      = 9
_BASE_CANVAS_FONT_L      = 10
_BASE_CANVAS_FONT_XL     = 11


def apply_display_scale(scale: float) -> None:
    """캔버스·레이아웃 상수를 디스플레이 스케일에 맞게 일괄 조정.
    constants 모듈 로드 시 자동 호출 → 다른 모듈이 'from constants import *' 로
    가져올 때 이미 스케일된 값을 얻는다.
    """
    global PX_PER_KM, RAMP_BAR_W, RAMP_BAR_H, RAMP_BAR_GAP_Y
    global RAMP_PX_PER_KM, RAMP_LANE_GAP_Y, IC_BOX_W, IC_GAP_MARGIN
    global CANVAS_BAR_H, CANVAS_TOP_MARGIN, CANVAS_DIR_GAP
    global CANVAS_KM_LBL_OFFSET, CANVAS_DIR_LBL_OFFSET, CANVAS_IC_CHAR_H
    global CANVAS_FONT_XS, CANVAS_FONT_S, CANVAS_FONT_M, CANVAS_FONT_L, CANVAS_FONT_XL
    global CANVAS_WIDGET_H, DETAIL_CANVAS_H, DISPLAY_SCALE

    s = max(0.5, min(2.0, float(scale)))
    DISPLAY_SCALE = s

    PX_PER_KM             = max(60,  round(_BASE_PX_PER_KM * s))
    CANVAS_WIDGET_H       = max(220, round(_BASE_CANVAS_WIDGET_H * s))
    DETAIL_CANVAS_H       = max(120, round(_BASE_DETAIL_CANVAS_H * s))
    RAMP_BAR_W            = max(200, round(_BASE_RAMP_BAR_W * s))
    RAMP_BAR_H            = max(10,  round(_BASE_RAMP_BAR_H * s))
    RAMP_BAR_GAP_Y        = max(15,  round(_BASE_RAMP_BAR_GAP_Y * s))
    RAMP_PX_PER_KM        = max(40,  round(_BASE_RAMP_PX_PER_KM * s))
    RAMP_LANE_GAP_Y       = max(3,   round(_BASE_RAMP_LANE_GAP_Y * s))
    IC_BOX_W              = max(20,  round(_BASE_IC_BOX_W * s))
    IC_GAP_MARGIN         = max(6,   round(_BASE_IC_GAP_MARGIN * s))
    CANVAS_BAR_H          = max(50,  round(_BASE_CANVAS_BAR_H * s))
    CANVAS_TOP_MARGIN     = max(30,  round(_BASE_CANVAS_TOP_MARGIN * s))
    CANVAS_DIR_GAP        = max(10,  round(_BASE_CANVAS_DIR_GAP * s))
    CANVAS_KM_LBL_OFFSET  = max(8,   round(_BASE_CANVAS_KM_LBL_OFF * s))
    CANVAS_DIR_LBL_OFFSET = max(18,  round(_BASE_CANVAS_DIR_LBL_OFF * s))
    CANVAS_IC_CHAR_H      = max(10,  round(_BASE_CANVAS_IC_CHAR_H * s))
    CANVAS_FONT_XS        = max(6,   round(_BASE_CANVAS_FONT_XS * s))
    CANVAS_FONT_S         = max(6,   round(_BASE_CANVAS_FONT_S * s))
    CANVAS_FONT_M         = max(7,   round(_BASE_CANVAS_FONT_M * s))
    CANVAS_FONT_L         = max(8,   round(_BASE_CANVAS_FONT_L * s))
    CANVAS_FONT_XL        = max(9,   round(_BASE_CANVAS_FONT_XL * s))


def _detect_auto_scale() -> float:
    """모니터 해상도로 기본 스케일을 자동 결정.
    window_mixin._auto_detect_scale 가 이 결과(DISPLAY_SCALE)를 그대로 사용하므로
    위젯 스케일과 캔버스(모식도) 스케일이 항상 일치한다.
    """
    try:
        import ctypes as _ct
        # ctk가 나중에 호출하는 것과 동일한 DPI 인식 활성화 → 물리 해상도 기준 측정
        try:
            _ct.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                _ct.windll.user32.SetProcessDPIAware()
            except Exception:
                pass
        sw = int(_ct.windll.user32.GetSystemMetrics(0))
    except Exception:
        return 1.0
    # 기준 해상도(REFERENCE_WIDTH=1920) 대비 정확한 비율로 스케일 결정
    #   예) 1280 → 1280/1920 = 0.667,  2560 → 1.333
    return max(0.5, min(2.0, sw / float(REFERENCE_WIDTH)))


def _init_display_scale_from_file() -> None:
    """앱 시작 시 display_settings.json 을 읽어 캔버스 상수를 자동 조정.
    constants 모듈이 처음 임포트될 때 실행되어, 이후 'from constants import *'
    하는 모든 모듈이 스케일된 값을 사용한다.
    파일이 없거나 'auto'면 모니터 해상도를 자동 감지해 적용한다.
    """
    scale = None
    try:
        import os as _os, json as _json, sys as _sys
        if getattr(_sys, 'frozen', False):
            base = _os.path.dirname(_sys.executable)
        else:
            base = _os.path.dirname(_os.path.abspath(__file__))
        path = _os.path.join(base, "display_settings.json")
        if _os.path.exists(path):
            with open(path, "r", encoding="utf-8") as _f:
                data = _json.load(_f)
            val = data.get("scale")
            if val and val != "auto":
                v = float(val)
                if 0.5 <= v <= 2.0:
                    scale = v
    except Exception:
        pass
    if scale is None:
        scale = _detect_auto_scale()
    try:
        apply_display_scale(scale)
    except Exception:
        pass


_init_display_scale_from_file()

# -------------------- 라이트 테마 --------------------
APP_BG = "#EEF4FB"
SURFACE_BG = "#F7FAFC"
CARD_BG = "#FFFFFF"
CARD_BORDER = "#D9E4F2"
NAV_BG = "#FFFFFF"
NAV_BORDER = "#D4E1F1"
NAV_TEXT = "#1F3A5F"
TITLE_TEXT = "#1B3C6B"
MUTED_TEXT = "#5F728C"
PRIMARY_BLUE = "#005BAC"
PRIMARY_BLUE_HOVER = "#004B8E"
ACCENT_RED = "#E33E4B"
ACCENT_RED_HOVER = "#C9313E"
APPLE_SHADOW = "#E6EEF8"

# -------------------- 유틸 --------------------
