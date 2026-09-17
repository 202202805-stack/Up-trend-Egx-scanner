
import sys
import subprocess

# تثبيت المكتبات المطلوبة تلقائياً في حال عدم وجودها
subprocess.check_call([sys.executable, "-m", "pip", "install", "yfinance", "scipy", "openpyxl"])

import os
import math
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import List, Literal, Optional, Tuple

import numpy as np
import pandas as pd
from scipy.signal import argrelextrema
import yfinance as yf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ======================================================================
# 1. CONFIGURATION
# ======================================================================

@dataclass
class CHSBConfig:
    """All tunable parameters for CHSB detector."""
    pivot_window: int = 5
    min_head_depth_pct: float = 0.015
    max_head_depth_pct: float = 0.60
    max_shoulder_price_asymmetry_pct: float = 0.05
    max_shoulder_time_asymmetry_pct: float = 0.35
    min_shoulder_above_head_pct: float = 0.01
    min_extra_shoulder_pairs: int = 1
    max_days_between_heads: int = 23
    max_head_price_asymmetry_pct: float = 0.15
    horizontal_neckline_slope_pct: float = 0.0006
    breakout_confirmation: Literal["close", "high"] = "close"
    max_bars_to_breakout: int = 90
    heavy_breakout_volume_multiple: float = 1.0
    volume_avg_window: int = 30
    throwback_window_days: int = 15
    throwback_tolerance_pct: float = 0.01
    stop_loss_buffer_pct: float = 0.0015
    stop_loss_buffer_abs: float = 0.0
    regime_sma_window: int = 100
    regime_lookback_slope_bars: int = 50


# ======================================================================
# 2. DATA STRUCTURES
# ======================================================================

@dataclass
class Pivot:
    idx: int
    date: pd.Timestamp
    price: float
    kind: Literal["low", "high"]


@dataclass
class ShoulderPair:
    left: Pivot
    right: Pivot
    price_asymmetry_pct: float
    time_asymmetry_pct: float


@dataclass
class NecklinePoint:
    idx: int
    date: pd.Timestamp
    price: float


@dataclass
class CHSBPattern:
    pattern_type: Literal["multiple_shoulders", "multiple_heads"]
    heads: List[Pivot]
    shoulder_pairs: List[ShoulderPair]
    neckline_left: NecklinePoint
    neckline_right: NecklinePoint
    neckline_slope_per_bar: float
    neckline_slope_class: Literal["up", "down", "horizontal"]

    formation_start_idx: int
    formation_end_idx: int

    breakout_idx: Optional[int]
    breakout_date: Optional[pd.Timestamp]
    breakout_price: Optional[float]
    breakout_is_gap: Optional[bool]
    breakout_volume_ratio: Optional[float]

    measure_rule_height: float
    measure_rule_target: float
    stop_loss: float

    throwback_occurred: Optional[bool]
    throwback_low: Optional[float]
    throwback_bars_to_resolve: Optional[int]

    volume_trend: Literal["rising", "falling", "flat", "unknown"]
    volume_left_avg: float
    volume_right_avg: float
    volume_shape: Literal["U", "dome", "random", "unknown"]

    market_regime: Literal["bull", "bear", "unknown"]
    width_days: int
    height_pct_of_price: float

    quality_score: float
    quality_notes: List[str] = field(default_factory=list)


# ======================================================================
# 3. PIVOT DETECTION
# ======================================================================

def _dedupe_extrema(idxs: np.ndarray, series: np.ndarray, minimize: bool) -> np.ndarray:
    if len(idxs) == 0:
        return idxs
    groups: List[List[int]] = [[int(idxs[0])]]
    for i in idxs[1:]:
        if i - groups[-1][-1] <= 1:
            groups[-1].append(int(i))
        else:
            groups.append([int(i)])
    out = []
    for g in groups:
        vals = series[g]
        best = g[int(np.argmin(vals))] if minimize else g[int(np.argmax(vals))]
        out.append(best)
    return np.array(out)


def find_pivots(df: pd.DataFrame, window: int) -> List[Pivot]:
    lows = df["low"].values
    highs = df["high"].values

    low_idx = argrelextrema(lows, np.less_equal, order=window)[0]
    high_idx = argrelextrema(highs, np.greater_equal, order=window)[0]

    low_idx = _dedupe_extrema(low_idx, lows, minimize=True)
    high_idx = _dedupe_extrema(high_idx, highs, minimize=False)

    pivots: List[Pivot] = []
    for i in low_idx:
        pivots.append(Pivot(idx=int(i), date=df.index[i], price=float(lows[i]), kind="low"))
    for i in high_idx:
        pivots.append(Pivot(idx=int(i), date=df.index[i], price=float(highs[i]), kind="high"))
    pivots.sort(key=lambda p: p.idx)
    return pivots


def pivot_lows(pivots: List[Pivot]) -> List[Pivot]:
    return [p for p in pivots if p.kind == "low"]


def pivot_highs(pivots: List[Pivot]) -> List[Pivot]:
    return [p for p in pivots if p.kind == "high"]


# ======================================================================
# 4. CORE STRUCTURAL DETECTION
# ======================================================================

def _pct_diff(a: float, b: float) -> float:
    denom = max(abs(a), abs(b), 1e-9)
    return abs(a - b) / denom


def _is_valid_shoulder_pair(left: Pivot, right: Pivot, head_price: float, cfg: CHSBConfig) -> Optional[ShoulderPair]:
    if left.price <= head_price * (1 + cfg.min_shoulder_above_head_pct):
        return None
    if right.price <= head_price * (1 + cfg.min_shoulder_above_head_pct):
        return None

    price_asym = _pct_diff(left.price, right.price)
    if price_asym > cfg.max_shoulder_price_asymmetry_pct:
        return None

    return ShoulderPair(left=left, right=right, price_asymmetry_pct=price_asym, time_asymmetry_pct=float("nan"))


def _time_asymmetry(head_idx: int, left_idx: int, right_idx: int) -> float:
    d_left = head_idx - left_idx
    d_right = right_idx - head_idx
    if d_left <= 0 or d_right <= 0:
        return float("inf")
    return abs(d_left - d_right) / max(d_left, d_right)


def _find_multiple_shoulder_patterns(lows: List[Pivot], highs: List[Pivot], head: Pivot, cfg: CHSBConfig) -> List[Tuple[List[ShoulderPair], int]]:
    left_candidates = sorted([p for p in lows if p.idx < head.idx], key=lambda p: -p.idx)
    right_candidates = sorted([p for p in lows if p.idx > head.idx], key=lambda p: p.idx)

    if not left_candidates or not right_candidates:
        return []

    chain: List[ShoulderPair] = []
    li, ri = 0, 0
    while li < len(left_candidates) and ri < len(right_candidates):
        L = left_candidates[li]
        R = right_candidates[ri]

        sp = _is_valid_shoulder_pair(L, R, head.price, cfg)
        if sp is None:
            if abs(L.idx - head.idx) < abs(R.idx - head.idx):
                ri += 1
            else:
                li += 1
            if li >= len(left_candidates) or ri >= len(right_candidates):
                break
            continue

        t_asym = _time_asymmetry(head.idx, L.idx, R.idx)
        if t_asym > cfg.max_shoulder_time_asymmetry_pct:
            li += 1
            ri += 1
            continue

        sp.time_asymmetry_pct = t_asym
        chain.append(sp)
        li += 1
        ri += 1

    if len(chain) < 1 + cfg.min_extra_shoulder_pairs:
        return []

    return [(chain, head.idx)]


def _find_multiple_head_patterns(lows: List[Pivot], cfg: CHSBConfig, head_candidates: Optional[List[Pivot]] = None) -> List[List[Pivot]]:
    base = head_candidates if head_candidates is not None else lows
    lows_sorted = sorted(base, key=lambda p: p.idx)
    clusters: List[List[Pivot]] = []
    used = set()

    for i, p in enumerate(lows_sorted):
        if p.idx in used:
            continue
        cluster = [p]
        for q in lows_sorted[i + 1:]:
            if q.idx - cluster[-1].idx > cfg.max_days_between_heads:
                break
            if _pct_diff(p.price, q.price) <= cfg.max_head_price_asymmetry_pct:
                cluster.append(q)
        if len(cluster) >= 2:
            for c in cluster:
                used.add(c.idx)
            clusters.append(cluster)
    return clusters


# ======================================================================
# 5. NECKLINE, BREAKOUT, MEASURE RULE, STOP LOSS
# ======================================================================

def _compute_neckline(highs: List[Pivot], left_bound_idx: int, center_idx: int, right_bound_idx: int) -> Optional[Tuple[NecklinePoint, NecklinePoint]]:
    left_peaks = [h for h in highs if left_bound_idx <= h.idx <= center_idx]
    right_peaks = [h for h in highs if center_idx <= h.idx <= right_bound_idx]
    if not left_peaks or not right_peaks:
        return None
    lp = max(left_peaks, key=lambda h: h.price)
    rp = max(right_peaks, key=lambda h: h.price)
    if lp.idx == rp.idx:
        return None
    return (NecklinePoint(lp.idx, lp.date, lp.price), NecklinePoint(rp.idx, rp.date, rp.price))


def _neckline_value_at(idx: int, nl_left: NecklinePoint, nl_right: NecklinePoint) -> float:
    if nl_right.idx == nl_left.idx:
        return nl_left.price
    slope = (nl_right.price - nl_left.price) / (nl_right.idx - nl_left.idx)
    return nl_left.price + slope * (idx - nl_left.idx)


def _classify_slope(nl_left: NecklinePoint, nl_right: NecklinePoint, cfg: CHSBConfig) -> str:
    if nl_right.idx == nl_left.idx:
        return "horizontal"
    slope = (nl_right.price - nl_left.price) / (nl_right.idx - nl_left.idx)
    avg_price = (nl_left.price + nl_right.price) / 2.0
    slope_pct = slope / max(avg_price, 1e-9)
    if abs(slope_pct) <= cfg.horizontal_neckline_slope_pct:
        return "horizontal"
    return "up" if slope_pct > 0 else "down"


def _find_breakout(df: pd.DataFrame, orig_df: pd.DataFrame, formation_end_idx: int, nl_left: NecklinePoint, nl_right: NecklinePoint, slope_class: str, head_low: float, right_shoulder_idx: int, cfg: CHSBConfig) -> dict:
    n = len(df)
    start = formation_end_idx + 1
    end = min(n, formation_end_idx + 1 + cfg.max_bars_to_breakout)

    if slope_class == "up":
        window = df.iloc[right_shoulder_idx if right_shoulder_idx < formation_end_idx else max(0, formation_end_idx - 1):formation_end_idx + 1]
        fixed_target = float(window["high"].max()) if len(window) else nl_right.price
    else:
        fixed_target = None

    price_col = "close" if cfg.breakout_confirmation == "close" else "high"

    for i in range(start, end):
        level = fixed_target if slope_class == "up" else _neckline_value_at(i, nl_left, nl_right)
        if df[price_col].iat[i] > level:
            is_gap = bool(df["low"].iat[i] > df["high"].iat[i - 1]) if i > 0 else False
            vol_avg_start = max(0, i - cfg.volume_avg_window)
            vol_avg = float(df["volume"].iloc[vol_avg_start:i].mean()) if i > vol_avg_start else np.nan
            vol_ratio = float(df["volume"].iat[i] / vol_avg) if vol_avg and not np.isnan(vol_avg) and vol_avg > 0 else np.nan

            return {
                "breakout_idx": i,
                "breakout_date": orig_df.index[i],
                "breakout_price": float(df[price_col].iat[i]),
                "breakout_level": float(level),
                "breakout_is_gap": is_gap,
                "breakout_volume_ratio": vol_ratio,
            }

    return {
        "breakout_idx": None, "breakout_date": None, "breakout_price": None,
        "breakout_level": None, "breakout_is_gap": None, "breakout_volume_ratio": None,
    }


def _measure_rule(nl_left: NecklinePoint, nl_right: NecklinePoint, head_low: float, breakout_price: float, head_idx: int) -> Tuple[float, float]:
    nl_at_head = _neckline_value_at(head_idx, nl_left, nl_right)
    height = max(nl_at_head - head_low, 0.0)
    target = breakout_price + height
    return height, target


def _stop_loss(lowest_point_price: float, cfg: CHSBConfig) -> float:
    buffer = lowest_point_price * cfg.stop_loss_buffer_pct + cfg.stop_loss_buffer_abs
    return lowest_point_price - buffer


def _detect_throwback(df: pd.DataFrame, breakout_idx: int, nl_left: NecklinePoint, nl_right: NecklinePoint, cfg: CHSBConfig) -> dict:
    n = len(df)
    end = min(n, breakout_idx + 1 + cfg.throwback_window_days)
    lowest, lowest_idx = None, None
    for i in range(breakout_idx + 1, end):
        nl_val = _neckline_value_at(i, nl_left, nl_right)
        low = df["low"].iat[i]
        if low <= nl_val * (1 + cfg.throwback_tolerance_pct):
            if lowest is None or low < lowest:
                lowest = float(low)
                lowest_idx = i
    if lowest is None:
        return {"throwback_occurred": False, "throwback_low": None, "throwback_bars_to_resolve": None}
    return {
        "throwback_occurred": True,
        "throwback_low": lowest,
        "throwback_bars_to_resolve": (lowest_idx - breakout_idx) if lowest_idx else None,
    }


# ======================================================================
# 6. VOLUME & REGIME ANALYSIS
# ======================================================================

def _volume_trend(df: pd.DataFrame, start_idx: int, end_idx: int) -> str:
    seg = df["volume"].iloc[start_idx:end_idx + 1].values
    if len(seg) < 3:
        return "unknown"
    x = np.arange(len(seg))
    slope = np.polyfit(x, seg, 1)[0]
    scale = np.mean(seg) if np.mean(seg) > 0 else 1.0
    slope_pct = slope / scale
    if slope_pct > 0.002: return "rising"
    if slope_pct < -0.002: return "falling"
    return "flat"


def _volume_shape(df: pd.DataFrame, start_idx: int, end_idx: int) -> str:
    seg = df["volume"].iloc[start_idx:end_idx + 1].values
    if len(seg) < 5:
        return "unknown"
    third = max(1, len(seg) // 3)
    left_avg = seg[:third].mean()
    mid_avg = seg[third:2 * third].mean()
    right_avg = seg[2 * third:].mean()
    if mid_avg > left_avg * 1.15 and mid_avg > right_avg * 1.15: return "U"
    if mid_avg < left_avg * 0.85 and mid_avg < right_avg * 0.85: return "dome"
    return "random"


def _market_regime(df: pd.DataFrame, at_idx: int, cfg: CHSBConfig) -> str:
    w = cfg.regime_sma_window
    if at_idx < w + cfg.regime_lookback_slope_bars:
        return "unknown"
    sma = df["close"].rolling(w).mean()
    now = sma.iat[at_idx]
    prev = sma.iat[at_idx - cfg.regime_lookback_slope_bars]
    if pd.isna(now) or pd.isna(prev):
        return "unknown"
    return "bull" if now > prev else "bear"


# ======================================================================
# 7. QUALITY SCORING
# ======================================================================

def score_pattern(p: CHSBPattern) -> Tuple[float, List[str]]:
    score = 50.0
    notes: List[str] = []

    if p.neckline_slope_class == "down": score += 12; notes.append("+12 down-sloping neckline")
    elif p.neckline_slope_class == "up": score += 3; notes.append("+3 up-sloping neckline")
    else: score += 1; notes.append("+1 horizontal neckline")

    if p.volume_trend == "rising": score += 10; notes.append("+10 rising volume trend")
    elif p.volume_trend == "falling": score += 2; notes.append("+2 falling volume trend")

    if p.volume_shape == "random": score += 6; notes.append("+6 random volume shape")
    elif p.volume_shape == "U": score += 4; notes.append("+4 U-shaped volume")

    if p.breakout_volume_ratio is not None and not math.isnan(p.breakout_volume_ratio):
        if p.breakout_volume_ratio >= 1.0: score += 8; notes.append("+8 heavy breakout volume")

    if p.throwback_occurred is False: score += 8; notes.append("+8 no throwback observed")
    elif p.throwback_occurred is True: score -= 4; notes.append("-4 throwback occurred")

    if p.breakout_is_gap: score += 6; notes.append("+6 gap on breakout day")

    if len(p.shoulder_pairs) >= 1 and p.shoulder_pairs[0].right.price > p.shoulder_pairs[0].left.price:
        score += 5; notes.append("+5 right shoulder low higher than left")

    if p.market_regime == "bull":
        if p.width_days >= 90: score += 4; notes.append("+4 wide pattern in bull market")
        if p.height_pct_of_price <= 0.1770: score += 3; notes.append("+3 short pattern in bull market")
        score += 3; notes.append("+3 bull-market regime")
    elif p.market_regime == "bear":
        if p.width_days <= 69: score += 4; notes.append("+4 narrow pattern in bear market")
        if p.height_pct_of_price >= 0.2784: score += 3; notes.append("+3 tall pattern in bear market")

    return max(0.0, min(100.0, score)), notes


# ======================================================================
# 8. TOP-LEVEL DETECTION PIPELINE
# ======================================================================

def _candidate_heads(lows: List[Pivot], cfg: CHSBConfig) -> List[Pivot]:
    out = []
    lows_sorted = sorted(lows, key=lambda p: p.idx)
    for i, p in enumerate(lows_sorted):
        neighbors = []
        if i > 0: neighbors.append(lows_sorted[i - 1])
        if i < len(lows_sorted) - 1: neighbors.append(lows_sorted[i + 1])
        if not neighbors: continue
        avg_neighbor = np.mean([n.price for n in neighbors])
        if avg_neighbor <= 0: continue
        depth = (avg_neighbor - p.price) / avg_neighbor
        if cfg.min_head_depth_pct <= depth <= cfg.max_head_depth_pct:
            out.append(p)
    return out


def _build_pattern_from_shoulder_chain(df: pd.DataFrame, orig_df: pd.DataFrame, heads: List[Pivot], chain: List[ShoulderPair], cfg: CHSBConfig, force_type: Optional[str] = None) -> Optional[CHSBPattern]:
    pattern_type = force_type or ("multiple_shoulders" if len(chain) >= 1 + cfg.min_extra_shoulder_pairs else None)
    if force_type is None and pattern_type is None: return None
    if force_type == "multiple_heads" and len(chain) < 1: return None

    outermost, innermost = chain[-1], chain[0]
    formation_start_idx, formation_end_idx = outermost.left.idx, outermost.right.idx
    center_idx = int(np.mean([h.idx for h in heads]))
    head_low = float(min(h.price for h in heads))
    head_idx_for_measure = min(heads, key=lambda h: h.price).idx

    highs_all = pivot_highs(find_pivots(df, cfg.pivot_window))
    nl = _compute_neckline(highs_all, formation_start_idx, center_idx, formation_end_idx)
    if nl is None: return None
    nl_left, nl_right = nl
    slope_class = _classify_slope(nl_left, nl_right, cfg)

    bo = _find_breakout(df, orig_df, formation_end_idx, nl_left, nl_right, slope_class, head_low, innermost.right.idx, cfg)

    if bo["breakout_price"] is not None:
        height, target = _measure_rule(nl_left, nl_right, head_low, bo["breakout_price"], head_idx_for_measure)
    else:
        nl_at_head = _neckline_value_at(head_idx_for_measure, nl_left, nl_right)
        height, target = max(nl_at_head - head_low, 0.0), float("nan")

    lowest_pts = [head_low] + [sp.left.price for sp in chain] + [sp.right.price for sp in chain]
    stop = _stop_loss(min(lowest_pts), cfg)

    tb = _detect_throwback(df, bo["breakout_idx"], nl_left, nl_right, cfg) if bo["breakout_idx"] is not None else {"throwback_occurred": None, "throwback_low": None, "throwback_bars_to_resolve": None}

    vt = _volume_trend(df, formation_start_idx, formation_end_idx)
    vshape = _volume_shape(df, formation_start_idx, formation_end_idx)
    vol_left = float(df["volume"].iloc[formation_start_idx:center_idx].mean()) if center_idx > formation_start_idx else float("nan")
    vol_right = float(df["volume"].iloc[center_idx:formation_end_idx + 1].mean()) if formation_end_idx >= center_idx else float("nan")

    regime = _market_regime(df, formation_start_idx, cfg)
    width_days = formation_end_idx - formation_start_idx
    avg_price_in_formation = float(df["close"].iloc[formation_start_idx:formation_end_idx + 1].mean())
    height_pct = height / avg_price_in_formation if avg_price_in_formation else float("nan")

    pat = CHSBPattern(
        pattern_type=pattern_type, heads=heads, shoulder_pairs=chain,
        neckline_left=nl_left, neckline_right=nl_right,
        neckline_slope_per_bar=(nl_right.price - nl_left.price) / max(1, (nl_right.idx - nl_left.idx)),
        neckline_slope_class=slope_class, formation_start_idx=formation_start_idx, formation_end_idx=formation_end_idx,
        breakout_idx=bo["breakout_idx"], breakout_date=bo["breakout_date"], breakout_price=bo["breakout_price"],
        breakout_is_gap=bo["breakout_is_gap"], breakout_volume_ratio=bo["breakout_volume_ratio"],
        measure_rule_height=height, measure_rule_target=target, stop_loss=stop,
        throwback_occurred=tb["throwback_occurred"], throwback_low=tb["throwback_low"],
        throwback_bars_to_resolve=tb["throwback_bars_to_resolve"], volume_trend=vt,
        volume_left_avg=vol_left, volume_right_avg=vol_right, volume_shape=vshape,
        market_regime=regime, width_days=width_days, height_pct_of_price=height_pct,
        quality_score=0.0, quality_notes=[],
    )
    pat.quality_score, pat.quality_notes = score_pattern(pat)
    return pat


def _dedupe_patterns(patterns: List[CHSBPattern]) -> List[CHSBPattern]:
    if not patterns: return patterns
    patterns_sorted = sorted(patterns, key=lambda p: (p.formation_start_idx, -p.quality_score))
    kept: List[CHSBPattern] = []
    for p in patterns_sorted:
        overlap = False
        for k in kept:
            if not (p.formation_end_idx < k.formation_start_idx or p.formation_start_idx > k.formation_end_idx):
                overlap = True
                if p.quality_score > k.quality_score:
                    kept.remove(k)
                    kept.append(p)
                break
        if not overlap: kept.append(p)
    kept.sort(key=lambda p: p.formation_start_idx)
    return kept


def detect_chsb(df: pd.DataFrame, cfg: Optional[CHSBConfig] = None) -> List[CHSBPattern]:
    if cfg is None: cfg = CHSBConfig()
    orig_df = df.copy()
    if not isinstance(orig_df.index, pd.DatetimeIndex):
        raise ValueError("df must be indexed by a DatetimeIndex")
    orig_df = orig_df.sort_index()
    
    df_indexed = orig_df.reset_index(drop=True)

    pivots = find_pivots(df_indexed, cfg.pivot_window)
    lows, highs = pivot_lows(pivots), pivot_highs(pivots)
    results: List[CHSBPattern] = []

    head_candidates = _candidate_heads(lows, cfg)
    for head in head_candidates:
        for chain, head_idx in _find_multiple_shoulder_patterns(lows, highs, head, cfg):
            pat = _build_pattern_from_shoulder_chain(df_indexed, orig_df, [head], chain, cfg)
            if pat is not None: results.append(pat)

    head_clusters = _find_multiple_head_patterns(lows, cfg, head_candidates=head_candidates)
    for cluster in head_clusters:
        cluster_sorted = sorted(cluster, key=lambda p: p.idx)
        left_idx, right_idx = cluster_sorted[0].idx, cluster_sorted[-1].idx
        left_shoulder_candidates = [p for p in lows if p.idx < left_idx]
        right_shoulder_candidates = [p for p in lows if p.idx > right_idx]
        if not left_shoulder_candidates or not right_shoulder_candidates: continue
        L = max(left_shoulder_candidates, key=lambda p: p.idx)
        R = min(right_shoulder_candidates, key=lambda p: p.idx)

        sp = _is_valid_shoulder_pair(L, R, float(np.mean([p.price for p in cluster_sorted])), cfg)
        if sp is None: continue
        sp.time_asymmetry_pct = _time_asymmetry(int(np.mean([left_idx, right_idx])), L.idx, R.idx)
        if sp.time_asymmetry_pct > cfg.max_shoulder_time_asymmetry_pct: continue

        pat = _build_pattern_from_shoulder_chain(df_indexed, orig_df, cluster_sorted, [sp], cfg, force_type="multiple_heads")
        if pat is not None: results.append(pat)

    return _dedupe_patterns(results)


# ======================================================================
# 9. BACKTESTING MAIN SCRIPT (1 YEAR WITH NON-OVERLAPPING TRADES RULE)
# ======================================================================

egyptian_stocks = [
    "AALR.CA", "ABUK.CA", "ACAMD.CA", "ACAP.CA", "ACGC.CA", "ACTF.CA", "ADCI.CA", "ADIB.CA",
    "ADPC.CA", "ADRI.CA", "AFDI.CA", "AFMC.CA", "AIDC.CA", "AIFI.CA", "AIH.CA", "AJWA.CA",
    "ALCN.CA", "ALEX.CA", "ALUM.CA", "AMER.CA", "AMES.CA", "AMIA.CA", "AMII.CA", "AMOC.CA",
    "AMPI.CA", "APSW.CA", "ARAB.CA", "ARCC.CA", "AREH.CA", "ASCM.CA", "ASPI.CA", "ATLC.CA",
    "ATQA.CA", "AXPH.CA", "BIDI.CA", "BIGP.CA", "BINV.CA", "BIOC.CA", "BONY.CA", "BTFH.CA",
    "CAED.CA", "CANA.CA", "CCAP.CA", "CCRS.CA", "CEFM.CA", "CERA.CA", "CFGH.CA", "CICH.CA",
    "CIEB.CA", "CIRA.CA", "CLHO.CA", "CNFN.CA", "COMI.CA", "COPR.CA", "COSG.CA", "CPCI.CA",
    "CPME.CA", "CRST.CA", "CSAG.CA", "DAPH.CA", "DCRC.CA", "DEIN.CA", "DGTZ.CA", "DOMT.CA",
    "DSCW.CA", "DTPP.CA", "EALR.CA", "EASB.CA", "EAST.CA", "EBSC.CA", "ECAP.CA", "EDFM.CA",
    "EEII.CA", "EFIC.CA", "EFID.CA", "EFIH.CA", "EGAL.CA", "EGAS.CA", "EGBE.CA", "EGCH.CA",
    "EGREF.CA", "EGSA.CA", "EGTS.CA", "EHDR.CA", "ELAB.CA", "ELEC.CA", "ELKA.CA", "ELNA.CA",
    "ELSH.CA", "ELWA.CA", "EMFD.CA", "ENGC.CA", "EOSB.CA", "EPCO.CA", "EPPK.CA", "ETEL.CA",
    "ETRS.CA", "EXPA.CA", "FAIT.CA", "FAITA.CA", "FCMD.CA", "FIRE.CA", "FNAR.CA", "FTNS.CA",
    "FWRY.CA", "GBCO.CA", "GDWA.CA", "GGCC.CA", "GGRN.CA", "GIHD.CA", "GMCI.CA", "GOUR.CA",
    "GPIM.CA", "GRCA.CA", "GSSC.CA", "GTEX.CA", "GTHE.CA", "GTWL.CA", "HBCO.CA", "HDBK.CA",
    "HELI.CA", "HRHO.CA", "IBCT.CA", "ICFC.CA", "ICID.CA", "IDRE.CA", "IEEC.CA", "IFAP.CA",
    "INEG.CA", "INFI.CA", "IRON.CA", "ISMA.CA", "ISMQ.CA", "ISPH.CA", "JUFO.CA", "KABO.CA",
    "KORA.CA", "KRDI.CA", "KWIN.CA", "KZPC.CA", "LCSW.CA", "LKGP.CA", "LUTS.CA", "MAAL.CA",
    "MASR.CA", "MBEG.CA", "MBSC.CA", "MCQE.CA", "MCRO.CA", "MENA.CA", "MEPA.CA", "MFPC.CA",
    "MFSC.CA", "MHOT.CA", "MICH.CA", "MILS.CA", "MIPH.CA", "MOED.CA", "MOIL.CA", "MOIN.CA",
    "MOSC.CA", "MPCI.CA", "MPCO.CA", "MPRC.CA", "MTIE.CA", "NAHO.CA", "NARE.CA", "NCCW.CA",
    "NCGC.CA", "NEDA.CA", "NHPS.CA", "NINH.CA", "NIPH.CA", "OBRI.CA", "OCAP.CA", "OCDI.CA",
    "OCPH.CA", "ODIN.CA", "OFH.CA", "OIH.CA", "OLFI.CA", "ORAS.CA", "ORHD.CA", "ORWE.CA",
    "PHAR.CA", "PHGC.CA", "PHTV.CA", "POUL.CA", "PRCL.CA", "PRDC.CA", "PRMH.CA", "QNBE.CA",
    "RACC.CA", "RAKT.CA", "RAYA.CA", "RKAZ.CA", "RMDA.CA", "RMTV.CA", "ROTO.CA", "RREI.CA",
    "RTVC.CA", "RUBX.CA", "SAUD.CA", "SCEM.CA", "SCFM.CA", "SCTS.CA", "SDTI.CA", "SEIG.CA",
    "SIEG.CA", "SIPC.CA", "SKPC.CA", "SMFR.CA", "SNFC.CA", "SPIN.CA", "SPMD.CA", "SUCE.CA",
    "SUGR.CA", "SVCE.CA", "SWDY.CA", "TALM.CA", "TANM.CA", "TAQA.CA", "TMGH.CA", "TORA.CA",
    "TWSA.CA", "TYCN.CA", "UBEE.CA", "UEFM.CA", "UEGC.CA", "UNIP.CA", "UNIT.CA", "UPMS.CA",
    "UTOP.CA", "VALU.CA", "VERT.CA", "VLMR.CA", "VLMRA.CA", "WCDF.CA", "WKOL.CA", "ZEOT.CA",
    "ZMID.CA",
]

end_date = datetime.now()
start_date = end_date - timedelta(days=1 * 365)

print(f"🚀 بدء تنفيذ الباك تست للفترة من {start_date.strftime('%Y-%m-%d')} إلى {end_date.strftime('%Y-%m-%d')} (سنة واحدة)...")

all_trades = []
open_trades = []
config = CHSBConfig()

for idx, ticker in enumerate(egyptian_stocks):
    print(f"[{idx+1}/{len(egyptian_stocks)}] جاري تنزيل وتحليل: {ticker}...")
    try:
        df = yf.download(ticker, start=start_date.strftime('%Y-%m-%d'), end=end_date.strftime('%Y-%m-%d'), progress=False)
        if df.empty or len(df) < 60:
            continue

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        df.columns = [c.lower() for c in df.columns]

        current_stock_price = float(df['close'].iloc[-1])
        patterns = detect_chsb(df, config)

        # ترتيب النماذج زمنياً لحساب تداخل الصفقات بشكل متسلسل
        valid_patterns = [p for p in patterns if p.breakout_idx is not None and p.breakout_price is not None and not np.isnan(p.measure_rule_target)]
        valid_patterns.sort(key=lambda x: x.breakout_idx)

        last_trade_exit_date = None

        for trade_idx, p in enumerate(valid_patterns):
            entry_idx = p.breakout_idx
            entry_date = df.index[entry_idx]

            # الشرط 1: إذا كان هناك صفقة قائمة على نفس السهم ولم تنتهِ بعد، نمنع فتح صفقة جديدة
            if last_trade_exit_date is not None and entry_date <= last_trade_exit_date:
                continue

            entry_price = p.breakout_price
            target_price = p.measure_rule_target
            stop_loss = p.stop_loss

            target_dist = target_price - entry_price
            pct_68_level = entry_price + (0.68 * target_dist)

            pivots_hist = find_pivots(df.reset_index(drop=True).iloc[:entry_idx], config.pivot_window)
            high_pivots = [pt.price for pt in pivots_hist if pt.kind == "high"]
            resistances_in_zone = [res for res in high_pivots if pct_68_level <= res < target_price]
            
            effective_target = min(resistances_in_zone) if resistances_in_zone else target_price

            # الشرط 2: أقصى مدة للصفقة 4 أشهر (120 يوماً) من تاريخ الفتح
            max_exit_date = entry_date + timedelta(days=120)

            exit_date = None
            exit_price = None
            status = "Open"

            for i in range(entry_idx + 1, len(df)):
                current_date = df.index[i]
                current_high = df['high'].iloc[i]
                current_low = df['low'].iloc[i]

                # تحقق الاستهداف أولاً
                if current_high >= effective_target:
                    exit_date = current_date
                    exit_price = effective_target
                    status = "WIN"
                    break

                # تحقق وقف الخسارة
                if current_low <= stop_loss:
                    exit_date = current_date
                    exit_price = stop_loss
                    status = "LOSS"
                    break

                # الإغلاق التلقائي بعد مرور 4 أشهر
                if current_date >= max_exit_date:
                    exit_date = current_date
                    exit_price = df['close'].iloc[i]
                    status = "WIN" if exit_price > entry_price else "LOSS"
                    break

            # تحديث تاريخ خروج آخر صفقة لمنع التداخل
            if exit_date is not None:
                last_trade_exit_date = exit_date
            else:
                # إذا كانت الصفقة مفتوحة حتى اليوم، نضع تاريخ الخروج كأقصى مدة (4 أشهر) لمنع فتح صفقات مجدداً
                last_trade_exit_date = max_exit_date

            trade_info = {
                "Stock Name": ticker,
                "Entry Date": entry_date.strftime('%Y-%m-%d'),
                "Entry Price": round(entry_price, 2),
                "Target": round(effective_target, 2),
                "Stop Loss": round(stop_loss, 2),
                "Exit Date": exit_date.strftime('%Y-%m-%d') if exit_date is not None else "-",
                "Current Price": round(current_stock_price, 2),
                "Status": status
            }

            all_trades.append(trade_info)

            if status == "Open":
                open_trades.append(trade_info)

    except Exception as e:
        print(f"❌ خطأ أثناء معالجة السهم {ticker}: {e}")

# ======================================================================
# 10. DISPLAY OPEN TRADES IN TERMINAL
# ======================================================================

print("\n" + "="*80)
print("📌 الصفقات المفتوحة (OPEN TRADES):")
print("="*80)
if open_trades:
    for ot in open_trades:
        print(f"🔹 السهم: {ot['Stock Name']} | تاريخ الدخول: {ot['Entry Date']} | سعر الدخول: {ot['Entry Price']} | الهدف: {ot['Target']} | وقف الخسارة: {ot['Stop Loss']} | السعر الحالي: {ot['Current Price']}")
else:
    print("لا توجد صفقات مفتوحة حالياً.")
print("="*80 + "\n")

# ======================================================================
# 11. EXPORTING STYLED EXCEL RESULTS
# ======================================================================

trades_df = pd.DataFrame(all_trades)

if not trades_df.empty:
    trades_df['Entry Date Temp'] = pd.to_datetime(trades_df['Entry Date'])
    trades_df = trades_df.sort_values(by='Entry Date Temp', ascending=True).reset_index(drop=True)
    trades_df = trades_df.drop(columns=['Entry Date Temp'])

    output_filename = "EGX_CHSB_Backtest_1Year.xlsx"
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "All Trades"

    headers = ["Stock Name", "Entry Date", "Entry Price", "Target", "Stop Loss", "Exit Date", "Current Price", "Status"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1B4373", end_color="1B4373", fill_type="solid") # أزرق داكن
    win_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")       # أخضر فاتح
    loss_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")      # أحمر/برتقالي فاتح
    open_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")      # أصفر فاتح

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=11)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, row_data in enumerate(trades_df.to_dict('records'), start=2):
        ws.append([row_data[h] for h in headers])
        status = row_data["Status"]

        if status == "WIN":
            row_fill = win_fill
        elif status == "LOSS":
            row_fill = loss_fill
        else:
            row_fill = open_fill

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = row_fill
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            if headers[col_idx - 1] in ["Entry Price", "Target", "Stop Loss", "Current Price"] and isinstance(cell.value, (int, float)):
                cell.number_format = '0.00'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    ws.freeze_panes = 'A2'
    wb.save(output_filename)

    print(f"✅ تم اكتمال الباك تست بنجاح!")
    print(f"📁 تم حفظ جميع الصفقات بالتنسيق المطلوب في ملف الإكسيل: {output_filename}")
else:
    print("\n⚠️ لم يتم العثور على أي صفقات مكتملة الشروط خلال الفترة المحددة.")
