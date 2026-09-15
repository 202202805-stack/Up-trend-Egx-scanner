import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import time
from dataclasses import dataclass, field, asdict
from typing import Optional
import numpy as np

try:
    from scipy.stats import linregress
    _HAS_SCIPY = True
except Exception:
    _HAS_SCIPY = False


# ===========================================================================
# Configuration -- Bulkowski Head-and-Shoulders Bottom Parameters
# ===========================================================================
@dataclass
class HSBConfig:
    bull_sma_period: int = 200
    atr_period: int = 14
    atr_multiplier: float = 3.0
    shoulder_match_pct: float = 2.0
    shoulder_reject_pct: float = 5.0
    head_min_depth_pct: float = 3.0
    time_symmetry_tol: float = 0.40
    min_formation_days: int = 15
    max_formation_days: int = 180
    prior_lookback_days: int = 60
    prior_decline_pct: float = 8.0
    min_r_squared: float = 0.40
    max_allowed_pct_per_bar: float = 0.003
    neckline_break_threshold: float = 0.0
    breakout_ma_period: int = 30
    stop_offset: float = 0.15
    target_mode: str = "measure_rule"
    confirmation_pct: float = 5.0
    horizon_days: int = 252
    throwback_window_days: int = 30
    median_height_pct_bull: float = 18.81
    median_height_pct_bear: float = 24.22
    median_length_days: int = 52


# ===========================================================================
# Helper Functions: Indicators & Trend Analysis
# ===========================================================================
def calculate_atr_wilder(df: pd.DataFrame, period: int = 14) -> pd.Series:
    high = df['high'] if 'high' in df.columns else df['High']
    low = df['low'] if 'low' in df.columns else df['Low']
    close = df['close'] if 'close' in df.columns else df['Close']
    close_prev = close.shift(1)
    
    tr1 = high - low
    tr2 = (high - close_prev).abs()
    tr3 = (low - close_prev).abs()
    
    tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
    atr = tr.ewm(alpha=1/period, adjust=False).mean()
    atr.iloc[:period-1] = np.nan
    return atr


def calculate_dynamic_atr_zigzag(df: pd.DataFrame, atr_period: int = 14, atr_multiplier: float = 3.0) -> list:
    df_calc = df.copy()
    df_calc['ATR'] = calculate_atr_wilder(df_calc, period=atr_period)
    
    high_col = 'high' if 'high' in df_calc.columns else 'High'
    low_col = 'low' if 'low' in df_calc.columns else 'Low'
    
    highs = df_calc[high_col].values
    lows = df_calc[low_col].values
    atrs = df_calc['ATR'].values
    
    pivots = []
    trend = 0  
    last_high_idx = 0
    last_low_idx = 0
    last_high = highs[0]
    last_low = lows[0]
    
    for i in range(atr_period, len(df_calc)):
        current_atr_threshold = atrs[i] * atr_multiplier
        if np.isnan(current_atr_threshold):
            continue
            
        if trend == 0:
            if highs[i] >= last_low + current_atr_threshold:
                trend = 1
                last_high, last_high_idx = highs[i], i
            elif lows[i] <= last_high - current_atr_threshold:
                trend = -1
                last_low, last_low_idx = lows[i], i
                
        elif trend == 1:
            if highs[i] > last_high:
                last_high, last_high_idx = highs[i], i
            elif lows[i] <= last_high - current_atr_threshold:
                pivots.append((last_high_idx, last_high, 'PEAK'))
                trend = -1
                last_low, last_low_idx = lows[i], i
                
        elif trend == -1:
            if lows[i] < last_low:
                last_low, last_low_idx = lows[i], i
            elif highs[i] >= last_low + current_atr_threshold:
                pivots.append((last_low_idx, last_low, 'TROUGH'))
                trend = 1
                last_high, last_high_idx = highs[i], i

    if trend == 1:
        pivots.append((last_high_idx, last_high, 'PEAK'))
    elif trend == -1:
        pivots.append((last_low_idx, last_low, 'TROUGH'))
                
    return pivots


def verify_prior_downtrend_advanced(closes: np.ndarray, ls_i: int, cfg: HSBConfig, high=None, low=None) -> bool:
    lookback = cfg.prior_lookback_days
    start_i = ls_i - lookback

    if start_i < 0 or ls_i > len(closes):
        return False

    window_closes = closes[start_i:ls_i]
    x = np.arange(lookback)
    
    slope, intercept = np.polyfit(x, window_closes, 1)
    normalized_slope = slope / window_closes[0]
    
    y_pred = slope * x + intercept
    residuals = window_closes - y_pred
    ss_res = np.sum(residuals**2)
    ss_tot = np.sum((window_closes - np.mean(window_closes))**2)
    r_squared = 1 - (ss_res / (ss_tot + 1e-9))

    total_decline = (window_closes[-1] - window_closes[0]) / window_closes[0]

    is_declining = total_decline <= -(cfg.prior_decline_pct / 100.0)
    is_steady_trend = (normalized_slope < 0) and (r_squared >= cfg.min_r_squared)

    return is_declining and is_steady_trend


# ===========================================================================
# Result Structure & Detector Class
# ===========================================================================
@dataclass
class HSBPattern:
    symbol: str
    left_shoulder_idx: int
    head_idx: int
    right_shoulder_idx: int
    ls_price: float
    head_price: float
    rs_price: float
    neckline_slope: float
    neckline_value_at_head: float
    p1_idx: int
    p2_idx: int
    p1_price: float
    p2_price: float
    formation_high: float
    formation_length_days: int
    height_pct: float
    grade: str
    shoulder_diff_pct: float
    time_symmetry_ratio: float
    downsloping_neckline: bool
    volume_trend: str
    volume_shape: str
    peak_volume_at: str
    breakout_volume: str
    market: str
    yearly_position: str
    entry: float
    stop_loss: float
    target: float
    risk_reward: float
    breakout_idx: Optional[int] = None
    breakout_date: Optional[str] = None
    exit_date: Optional[str] = None
    current_price: Optional[float] = None
    ls_date: Optional[str] = None
    head_date: Optional[str] = None
    rs_date: Optional[str] = None
    breakout_gap: bool = False
    status: str = "forming"
    hit_target: Optional[bool] = None
    max_rise_pct: Optional[float] = None
    max_loss_pct: Optional[float] = None
    days_to_ultimate_high: Optional[int] = None
    change_after_trend_end: Optional[float] = None
    throwback: Optional[bool] = None
    quality_score: float = 0.0
    reasons: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)


class HSBottomDetector:
    def __init__(self, config: Optional[HSBConfig] = None):
        self.cfg = config or HSBConfig()

    @staticmethod
    def _neckline_at(p1_i, p1_p, p2_i, p2_p, t):
        slope = (p2_p - p1_p) / max(p2_i - p1_i, 1)
        return p1_p + slope * (t - p1_i), slope

    def _volume_diag(self, df, ls_i, h_i, rs_i, p1_i, p2_i):
        vol_col = 'volume' if 'volume' in df.columns else 'Volume'
        vol = df[vol_col].values
        seg = vol[ls_i:rs_i + 1]
        n = len(seg)
        if _HAS_SCIPY and n > 5:
            slope = linregress(np.arange(n), seg).slope
            vol_trend = "falling" if slope < 0 else "rising"
        else:
            half = max(n // 2, 1)
            vol_trend = "falling" if seg[:half].mean() > seg[-half:].mean() else "rising"
        t = max(n // 3, 1)
        outer = np.concatenate([seg[:t], seg[-t:]]).mean()
        middle = seg[t:-t].mean() if n > 2 * t else seg.mean()
        vol_shape = "U" if outer > middle * 1.05 else ("dome" if middle > outer * 1.05 else "none")

        def _around(i):
            return vol[max(i - 5, 0): i + 6].mean()
        avgs = {"LS": _around(ls_i), "HEAD": _around(h_i), "RS": _around(rs_i)}
        return vol_trend, vol_shape, max(avgs, key=avgs.get)

    def detect(self, df: pd.DataFrame, symbol: str = "") -> list:
        cfg = self.cfg
        df = df.reset_index(drop=True)
        close_col = 'close' if 'close' in df.columns else 'Close'
        high_col = 'high' if 'high' in df.columns else 'High'
        low_col = 'low' if 'low' in df.columns else 'Low'
        
        closes = df[close_col].values
        highs = df[high_col].values
        lows = df[low_col].values
        sma = df[close_col].rolling(cfg.bull_sma_period).mean().values

        pivots = calculate_dynamic_atr_zigzag(df, atr_period=cfg.atr_period, atr_multiplier=cfg.atr_multiplier)
        patterns = []
        for i in range(len(pivots) - 4):
            (ls_i, ls_p, k0), (p1_i, p1_p, k1), (h_i, h_p, k2), \
                (p2_i, p2_p, k3), (rs_i, rs_p, k4) = pivots[i:i + 5]
                
            if not (k0 == 'TROUGH' and k1 == 'PEAK' and k2 == 'TROUGH' and k3 == 'PEAK' and k4 == 'TROUGH'):
                continue
                
            pat = self._validate(df, closes, highs, lows, sma, symbol,
                                 ls_i, ls_p, p1_i, p1_p, h_i, h_p, p2_i, p2_p, rs_i, rs_p)
            if pat is not None:
                patterns.append(pat)
                
        return self._dedup(patterns)

    def _validate(self, df, closes, highs, lows, sma, symbol,
                  ls_i, ls_p, p1_i, p1_p, h_i, h_p, p2_i, p2_p, rs_i, rs_p):
        cfg = self.cfg
        length = rs_i - ls_i
        if not (cfg.min_formation_days <= length <= cfg.max_formation_days):
            return None

        if not verify_prior_downtrend_advanced(closes, ls_i, cfg, high=highs, low=lows):
            return None

        shoulder_ref = min(ls_p, rs_p)
        if not (h_p < ls_p and h_p < rs_p):
            return None
        if ((shoulder_ref - h_p) / shoulder_ref * 100.0) < cfg.head_min_depth_pct:
            return None

        shoulder_diff = abs(ls_p - rs_p) / ((ls_p + rs_p) / 2.0) * 100.0
        if shoulder_diff > cfg.shoulder_reject_pct:
            return None
        grade = "A" if shoulder_diff <= cfg.shoulder_match_pct else "B"

        dt1, dt2 = h_i - ls_i, rs_i - h_i
        if min(dt1, dt2) < 3 or (min(dt1, dt2) / max(dt1, dt2)) < (1.0 - cfg.time_symmetry_tol):
            return None

        bar_distance = p2_i - p1_i
        if bar_distance <= 0 or abs(((p2_p - p1_p) / p1_p) / bar_distance) > cfg.max_allowed_pct_per_bar:
            return None

        nl_head, slope = self._neckline_at(p1_i, p1_p, p2_i, p2_p, h_i)
        formation_high = max(highs[p1_i:p2_i + 1])
        head_low = lows[max(0, h_i - 3): min(len(lows), h_i + 4)].min()
        ls_low = lows[max(0, ls_i - 3): min(len(lows), ls_i + 4)].min()
        rs_low = lows[max(0, rs_i - 3): min(len(lows), rs_i + 4)].min()

        height_pct = (nl_head - head_low) / formation_high * 100.0
        vol_trend, vol_shape, peak_at = self._volume_diag(df, ls_i, h_i, rs_i, p1_i, p2_i)

        mkt = "bull" if sma[rs_i] == sma[rs_i] and closes[rs_i] > sma[rs_i] else "bear"
        yr_hi = highs[max(rs_i - 252, 0): rs_i + 1].max()
        yr_lo = lows[max(rs_i - 252, 0): rs_i + 1].min()
        pos = (closes[rs_i] - yr_lo) / max(yr_hi - yr_lo, 1e-9)
        yearly_pos = "low" if pos < 1 / 3 else ("high" if pos > 2 / 3 else "center")

        entry = closes[rs_i]
        stop = min(ls_low, rs_low) - cfg.stop_offset
        target = formation_high + (nl_head - head_low)
        risk = entry - stop

        ls_date = str(df['date'].iloc[ls_i]) if 'date' in df.columns else str(ls_i)
        head_date = str(df['date'].iloc[h_i]) if 'date' in df.columns else str(h_i)
        rs_date = str(df['date'].iloc[rs_i]) if 'date' in df.columns else str(rs_i)

        pat = HSBPattern(
            symbol=symbol, left_shoulder_idx=ls_i, head_idx=h_i, right_shoulder_idx=rs_i,
            ls_price=ls_p, head_price=h_p, rs_price=rs_p, neckline_slope=slope,
            neckline_value_at_head=nl_head, p1_idx=p1_i, p2_idx=p2_i, p1_price=p1_p, p2_price=p2_p,
            formation_high=formation_high, formation_length_days=length, height_pct=height_pct,
            grade=grade, shoulder_diff_pct=shoulder_diff, time_symmetry_ratio=min(dt1, dt2) / max(dt1, dt2),
            downsloping_neckline=slope < 0, volume_trend=vol_trend, volume_shape=vol_shape,
            peak_volume_at=peak_at, breakout_volume="", market=mkt, yearly_position=yearly_pos,
            entry=entry, stop_loss=stop, target=target, risk_reward=(target - entry) / risk if risk > 0 else np.nan,
            ls_date=ls_date, head_date=head_date, rs_date=rs_date
        )
        self._score(pat)
        return pat

    def _score(self, pat: HSBPattern) -> None:
        cfg = self.cfg
        s, why = 50.0, []
        med_h = cfg.median_height_pct_bull if pat.market == "bull" else cfg.median_height_pct_bear

        if pat.grade == "A": s += 12; why.append("shoulders within 2% (Grade A)")
        else: s += 6; why.append("shoulders within 5% (Grade B)")
        
        if pat.time_symmetry_ratio >= 0.8: s += 8; why.append("strong time symmetry")
        if pat.height_pct >= med_h: s += 8; why.append(f"tall pattern ({pat.height_pct:.1f}% >= {med_h}%)")
        if pat.formation_length_days <= cfg.median_length_days: s += 8; why.append("narrow pattern")
        if pat.volume_trend == "falling": s += 6; why.append("falling volume trend")
        if pat.volume_shape == "U": s += 6; why.append("U-shaped volume")
        if pat.market == "bull": s += 10; why.append("bull market")
        if pat.market == "bull" and pat.yearly_position == "high": s += 8; why.append("breakout near yearly high")
        if pat.market == "bear" and pat.yearly_position == "low": s += 8; why.append("breakout near yearly low")
        if pat.market == "bull" and pat.downsloping_neckline: s += 6; why.append("down-sloping neckline")
        if pat.ls_price > pat.rs_price: s += 6; why.append("higher left-shoulder low")

        pat.quality_score = round(min(s, 100.0), 1)
        pat.reasons = why

    def evaluate(self, df: pd.DataFrame, patterns: list) -> list:
        cfg = self.cfg
        df = df.reset_index(drop=True)
        close_col = 'close' if 'close' in df.columns else 'Close'
        high_col = 'high' if 'high' in df.columns else 'High'
        low_col = 'low' if 'low' in df.columns else 'Low'
        open_col = 'open' if 'open' in df.columns else 'Open'
        vol_col = 'volume' if 'volume' in df.columns else 'Volume'
        
        closes = df[close_col].values
        highs = df[high_col].values
        lows = df[low_col].values
        vol = df[vol_col].values
        vol_ma = df[vol_col].rolling(cfg.breakout_ma_period).mean().values
        latest_close = closes[-1]

        for pat in patterns:
            p1_idx, p1_price = pat.p1_idx, pat.p1_price
            slope = (pat.p2_price - p1_price) / max(pat.p2_idx - p1_idx, 1)
            bo_i = None
            start = pat.right_shoulder_idx + 1
            
            for t in range(start, len(df)):
                nl_t = p1_price + slope * (t - p1_idx)
                if closes[t] > nl_t * (1 + cfg.neckline_break_threshold):
                    bo_i = t
                    break
                    
            if bo_i is None:
                pat.status = "forming"
                continue

            pat.breakout_idx = bo_i
            pat.breakout_date = str(df['date'].iloc[bo_i]) if 'date' in df.columns else str(bo_i)
            pat.entry = closes[bo_i]
            pat.current_price = latest_close
            pat.breakout_gap = df[open_col].values[bo_i] > highs[bo_i - 1]
            pat.breakout_volume = "heavy" if vol[bo_i] > vol_ma[bo_i] else "light"

            pat.target = pat.entry + (pat.neckline_value_at_head - pat.head_price)
            risk = pat.entry - pat.stop_loss
            pat.risk_reward = (pat.target - pat.entry) / risk if risk > 0 else np.nan

            # تتبع الصفقة حتى الهدف أو وقف الخسارة أو استمرار فتحها
            hit_win = False
            hit_loss = False
            exit_idx = None

            for t in range(bo_i + 1, len(df)):
                if highs[t] >= pat.target:
                    hit_win = True
                    exit_idx = t
                    break
                elif lows[t] <= pat.stop_loss:
                    hit_loss = True
                    exit_idx = t
                    break

            if hit_win:
                pat.status = "WIN"
                pat.hit_target = True
                pat.exit_date = str(df['date'].iloc[exit_idx]) if 'date' in df.columns else str(exit_idx)
            elif hit_loss:
                pat.status = "LOSS"
                pat.hit_target = False
                pat.exit_date = str(df['date'].iloc[exit_idx]) if 'date' in df.columns else str(exit_idx)
            else:
                pat.status = "OPEN"
                pat.hit_target = None
                pat.exit_date = ""  # يبقى فارغًا للصفقات المفتوحة

        return patterns

    def _dedup(self, patterns: list) -> list:
        seen = {}
        for p in patterns:
            key = p.head_idx
            if key not in seen or (p.grade, p.quality_score) > (seen[key].grade, seen[key].quality_score):
                seen[key] = p
        return list(seen.values())


# ===========================================================================
# Backtest Runner
# ===========================================================================
def run_egx_backtest(symbols_list):
    config = HSBConfig()
    detector = HSBottomDetector(config)
    all_trades = []
    
    # باك تست لمدة سنة واحدة فقط (365 يوم)
    end_date = datetime.today().strftime('%Y-%m-%d')
    start_date = (datetime.today() - timedelta(days=365)).strftime('%Y-%m-%d')
    
    print(f"=== بدء فحص الباك تست لمدة سنة واحدة لعدد {len(symbols_list)} سهم ===")
    
    for symbol in symbols_list:
        try:
            print(f"جاري جلب وفحص بيانات السهم: {symbol}...")
            df = yf.download(symbol, start=start_date, end=end_date, progress=False)
            
            if df.empty:
                print(f"  -> لم يتم العثور على بيانات للسهم {symbol}. تخطي.")
                continue

            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)
            
            df = df.reset_index()
            df.columns = [c.lower() for c in df.columns]
            
            if 'date' in df.columns:
                df['date'] = pd.to_datetime(df['date']).dt.strftime('%Y-%m-%d')
            
            if len(df) < 60:
                print(f"  -> بيانات غير كافية للسهم {symbol}. تخطي.")
                continue
                
            patterns = detector.detect(df, symbol=symbol)
            detector.evaluate(df, patterns)
            
            for pat in patterns:
                # تضمين الصفقات المغلقة (WIN / LOSS) والصفقات المفتوحة (OPEN)
                if pat.status not in ["WIN", "LOSS", "OPEN"]:
                    continue

                trade_record = {
                    "Stock Name": pat.symbol,
                    "Entry Date": pat.breakout_date,
                    "Entry Price": round(pat.entry, 3),
                    "Target": round(pat.target, 3),
                    "Stop Loss": round(pat.stop_loss, 3),
                    "Exit Date": pat.exit_date if pat.exit_date else "",
                    "Current Price": round(pat.current_price, 3) if pat.current_price is not None else None,
                    "Status": pat.status
                }
                all_trades.append(trade_record)
                
            time.sleep(0.2)
            
        except Exception as e:
            print(f"  -> حدث خطأ أثناء معالجة السهم {symbol}: {e}")
            
    if all_trades:
        results_df = pd.DataFrame(all_trades)
        
        # ترتيب الصفقات من الأقدم إلى الأحدث حسب تاريخ الدخول (Entry Date)
        results_df['Entry Date'] = pd.to_datetime(results_df['Entry Date'])
        results_df = results_df.sort_values(by='Entry Date', ascending=True).reset_index(drop=True)
        results_df['Entry Date'] = results_df['Entry Date'].dt.strftime('%Y-%m-%d')

        output_filename = "head_and_shoulders_bottom_results.xlsx"
        
        # حفظ النتائج في صفحة واحدة تحت اسم All Trades List
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='All Trades List', index=False)

        print(f"\n[تم بنجاح] تم حفظ جميع الصفقات والنتائج في الملف: {output_filename}")
        print(f"إجمالي عدد الصفقات المسجلة: {len(results_df)}")
        return results_df
    else:
        print("\n[تنبيه] لم يتم العثور على صفقات مطابقة بالشروط المحددة.")
        return pd.DataFrame()


# ===========================================================================
# Execution Section
# ===========================================================================
if __name__ == "__main__":
    symbols = [
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
        "PHAR.CA", "PHDC.CA", "PHGC.CA", "PHTV.CA", "POUL.CA", "PRCL.CA", "PRDC.CA", "PRMH.CA",
        "QNBE.CA", "RACC.CA", "RAKT.CA", "RAYA.CA", "RKAZ.CA", "RMDA.CA", "RMTV.CA", "ROTO.CA",
        "RREI.CA", "RTVC.CA", "RUBX.CA", "SAUD.CA", "SCEM.CA", "SCFM.CA", "SCTS.CA", "SDTI.CA",
        "SEIG.CA", "SIEG.CA", "SIPC.CA", "SKPC.CA", "SMFR.CA", "SNFC.CA", "SPIN.CA", "SPMD.CA",
        "SUCE.CA", "SUGR.CA", "SVCE.CA", "SWDY.CA", "TALM.CA", "TANM.CA", "TAQA.CA", "TMGH.CA",
        "TORA.CA", "TWSA.CA", "TYCN.CA", "UBEE.CA", "UEFM.CA", "UEGC.CA", "UNIP.CA", "UNIT.CA",
        "UPMS.CA", "UTOP.CA", "VALU.CA", "VERT.CA", "VLMR.CA", "VLMRA.CA", "WCDF.CA", "WKOL.CA",
        "ZEOT.CA", "ZMID.CA"
    ]

    results = run_egx_backtest(symbols)
