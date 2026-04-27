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
