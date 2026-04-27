# -*- coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 - 캔버스 렌더링 유틸리티
"""
import math
import tkinter as tk
from constants import *
from utils import get_method_warranty_period, intervals_overlap


# -------------------- 캔버스 렌더링 모듈 유틸 --------------------
def _split_gaps(gaps_px, xl, xr):
    """gaps_px 내 IC 갭을 제외한 [xl, xr] 구간의 서브 세그먼트 목록 반환."""
    segments = []
    cur = xl
    for a, b in gaps_px:
        if b <= cur:
            continue
        if a >= xr:
            break
        if a > cur:
            segments.append((cur, min(a, xr)))
        cur = max(cur, b)
    if cur < xr:
        segments.append((cur, xr))
    return segments


def _merge_gaps(raw_gaps):
    """겹치는 갭 구간을 병합하여 정렬된 (left, right) 리스트 반환."""
    if not raw_gaps:
        return []
    result = [list(raw_gaps[0])]
    for a, b in raw_gaps[1:]:
        if a > result[-1][1] + 1:
            result.append([a, b])
        else:
            result[-1][1] = max(result[-1][1], b)
    return [(a, b) for a, b in result]


def _draw_hatch_rect(canvas, xl, yt, xr, yb, color="#E53E3E"):
    """직사각형 내부에 대각선 빗금을 그린다."""
    if xr - xl <= 0 or yb - yt <= 0:
        return
    spacing = 8
    height  = yb - yt
    t = -height
    while t <= (xr - xl):
        xs = max(xl + t,          xl)
        xe = min(xl + t + height, xr)
        if xe > xs:
            ys = yt + (xs - (xl + t))
            ye = yt + (xe - (xl + t))
            canvas.create_line(xs, ys, xe, ye, fill=color, width=1)
        t += spacing


def _compute_overlap_status(visible_entries):
    """
    하자기간 기반 중복 상태 계산 (chain 방식).

    Returns: {idx: 'normal' | 'overlap' | 'hidden'}
      - 'normal'  : 정상 표시
      - 'overlap' : 하자기간 내 중복 시공 (빗금 표시)
      - 'hidden'  : 하자기간 만료 후 신규 보수로 대체됨 (숨김, 데이터는 보존)

    규칙:
      - 새 보수의 시공연도 ≤ 앵커연도 + 하자기간  → 중복(overlap)
      - 새 보수의 시공연도 > 앵커연도 + 하자기간  → 새 앵커로 교체, 이전 앵커 체인 전체 숨김
      - 하자기간은 카테고리별 설정값 우선, 기본값: 표면개량=2년, 그 외=3년
    """
    entries_data = []
    for idx, it, year in visible_entries:
        try:
            yr = int(year) if year else None
        except (ValueError, TypeError):
            yr = None
        entries_data.append({
            'idx': idx,
            'entry': it,
            'year': yr,
            'warranty': get_method_warranty_period(it.get("method", "")),
            'start': float(it.get('start', 0)),
            'end': float(it.get('end', 0)),
            'dir': str(it.get('direction', '')),
            'lane': str(it.get('lane', '전차로')),
        })

    status = {d['idx']: 'normal' for d in entries_data}

    # (방향, 차로) 별로 그룹화
    groups: dict = {}
    for d in entries_data:
        key = (d['dir'], d['lane'])
        groups.setdefault(key, []).append(d)

    for key, group in groups.items():
        # 연도 오름차순 정렬 (연도 없는 항목은 마지막)
        group.sort(key=lambda x: (x['year'] is None, x['year'] or 0))

        # active_anchors: 현재 유효한 앵커 목록
        # 각 항목: {'d': entry_data, 'members': [idx, ...]}
        active_anchors: list = []

        for d in group:
            yr = d['year']
            if yr is None:
                continue  # 연도 정보 없으면 기본(normal) 유지

            # 공간적으로 겹치는 앵커 찾기
            overlapping = [a for a in active_anchors
                           if intervals_overlap(d['start'], d['end'],
                                               a['d']['start'], a['d']['end'])]

            if not overlapping:
                # 겹치는 앵커 없음 → 새 앵커
                active_anchors.append({'d': d, 'members': [d['idx']]})
                status[d['idx']] = 'normal'
                continue

            # 하자기간 내 앵커 vs 만료된 앵커 분류
            in_warranty = [a for a in overlapping
                           if yr <= a['d']['year'] + a['d']['warranty']]
            expired     = [a for a in overlapping
                           if yr > a['d']['year'] + a['d']['warranty']]

            if in_warranty:
                # 하자기간 내 → 중복 표기
                status[d['idx']] = 'overlap'
                # 가장 최근 유효 앵커의 체인에 추가 (앵커 자체가 되지는 않음)
                latest_anchor = max(in_warranty, key=lambda a: a['d']['year'])
                latest_anchor['members'].append(d['idx'])
            else:
                # 모든 겹치는 앵커가 만료됨 → 새 앵커로 대체
                # 만료된 앵커와 그 체인 전체를 숨김 처리
                for a in expired:
                    for member_idx in a['members']:
                        status[member_idx] = 'hidden'
                    active_anchors.remove(a)
                # 새 앵커 등록
                active_anchors.append({'d': d, 'members': [d['idx']]})
                status[d['idx']] = 'normal'

    return status
