import datetime
import os
import time
import warnings
from dataclasses import dataclass
from typing import List, Optional

import numpy as np
import openpyxl
import pandas as pd
import requests
import yfinance as yf
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from scipy.signal import argrelextrema

warnings.filterwarnings('ignore')


# ==============================================================================
# 1. TRADINGVIEW INTEGRATION FOR LIVE DATA
# ==============================================================================

def fetch_tradingview_live_data(tickers: List[str]) -> dict:
    """جلب الأسعار اللحظية عبر TradingView"""
    tv_data = {}
    symbols = [f"EGX:{ticker.replace('.CA', '')}" for ticker in tickers]
    
    url = "https://scanner.tradingview.com/egypt/scan"
    payload = {
        "symbols": {"tickers": symbols},
        "columns": ["name", "open", "high", "low", "close", "volume"]
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    try:
        response = requests.post(url, json=payload, headers=headers, timeout=15)
        if response.status_code == 200:
            res_json = response.json()
            for item in res_json.get("data", []):
                raw_ticker = item.get("s", "")
                ticker = raw_ticker.replace("EGX:", "") + ".CA"
                vals = item.get("d", [])
                if len(vals) >= 6 and vals[1] is not None:
                    tv_data[ticker] = {
                        "Open": float(vals[1]),
                        "High": float(vals[2]),
                        "Low": float(vals[3]),
                        "Close": float(vals[4]),
                        "Volume": float(vals[5]) if vals[5] is not None else 0.0
                    }
    except Exception as e:
        print(f"    [TV Warning] Could not fetch live data from TradingView: {e}")

    return tv_data


def get_combined_stock_data(ticker: str, start_date: datetime.datetime, end_date: datetime.datetime, tv_live_dict: dict) -> pd.DataFrame:
    """تحميل بيانات Yahoo Finance ودمج آخر سعر لحظي من TradingView مع مراعاة توقيت مصر"""
    try:
        df = yf.download(
            ticker,
            start=start_date.strftime('%Y-%m-%d'),
            end=end_date.strftime('%Y-%m-%d'),
            progress=False
        )
    except Exception:
        return pd.DataFrame()

    if df.empty:
        return pd.DataFrame()

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    df.columns = [c.lower() for c in df.columns]

    if ticker in tv_live_dict:
        tv_bar = tv_live_dict[ticker]
        # استخدام التوقيت المحلي لتفادي فروق التوقيت في GitHub Actions
        today_date = pd.Timestamp(datetime.datetime.now().date())
        
        if df.index[-1].date() < today_date.date():
            new_row = pd.DataFrame({
                'open': [tv_bar['Open']],
                'high': [tv_bar['High']],
                'low': [tv_bar['Low']],
                'close': [tv_bar['Close']],
                'volume': [tv_bar['Volume']]
            }, index=[today_date])
            df = pd.concat([df, new_row])
        else:
            df.iloc[-1, df.columns.get_loc('high')] = max(df.iloc[-1]['high'], tv_bar['High'])
            df.iloc[-1, df.columns.get_loc('low')] = min(df.iloc[-1]['low'], tv_bar['Low'])
            df.iloc[-1, df.columns.get_loc('close')] = tv_bar['Close']
            df.iloc[-1, df.columns.get_loc('volume')] = max(df.iloc[-1]['volume'], tv_bar['Volume'])

    df.ffill(inplace=True)
    df.bfill(inplace=True)
    return df


# ==============================================================================
# 2. DATA CLASSES & BROADENING BOTTOM DETECTOR
# ==============================================================================

@dataclass
class SwingPoint:
    index: int
    date: pd.Timestamp
    price: float
    type: str

@dataclass
class PerfectBroadeningBottom:
    start_date: pd.Timestamp
    entry_date: pd.Timestamp
    p1: SwingPoint
    p2: SwingPoint
    p3: SwingPoint
    p4: SwingPoint
    p5: SwingPoint
    upper_slope: float
    upper_intercept: float
    lower_slope: float
    lower_intercept: float
    confidence: float

    def upper_value_at(self, x: float) -> float:
        return self.upper_slope * x + self.upper_intercept

    def lower_value_at(self, x: float) -> float:
        return self.lower_slope * x + self.lower_intercept


class BroadeningBottomDetector:
    def __init__(self, order: int = 4, max_linearity_error: float = 0.020): # تم رفع النسبة قليلاً لتطابق نتائج VS Code
        self.order = order
        self.max_linearity_error = max_linearity_error

    def find_swing_points(self, df_window: pd.DataFrame) -> List[SwingPoint]:
        highs = df_window['high'].values
        lows = df_window['low'].values
        dates = df_window.index

        maxima_idx = argrelextrema(highs, np.greater, order=self.order)[0]
        minima_idx = argrelextrema(lows, np.less, order=self.order)[0]

        swing_points = []
        for idx in maxima_idx:
            swing_points.append(SwingPoint(int(idx), dates[idx], float(highs[idx]), 'high'))
        for idx in minima_idx:
            swing_points.append(SwingPoint(int(idx), dates[idx], float(lows[idx]), 'low'))

        swing_points.sort(key=lambda x: x.index)

        filtered = []
        for sp in swing_points:
            if not filtered or sp.type != filtered[-1].type:
                filtered.append(sp)
            elif sp.type == 'high' and sp.price > filtered[-1].price:
                filtered[-1] = sp
            elif sp.type == 'low' and sp.price < filtered[-1].price:
                filtered[-1] = sp

        return filtered

    def detect(self, df_window: pd.DataFrame) -> Optional[PerfectBroadeningBottom]:
        swings = self.find_swing_points(df_window)
        
        if len(swings) < 5:
            return None

        best_pattern = None
        best_confidence = 0

        for i in range(len(swings) - 4):
            p1, p2, p3, p4, p5 = swings[i : i + 5]

            if not (p1.type == 'low' and p2.type == 'high' and p3.type == 'low' and p4.type == 'high' and p5.type == 'low'):
                continue

            p1_abs_loc = df_window.index.get_loc(p1.date)
            if p1_abs_loc >= 10:
                prior_data = df_window.iloc[:p1_abs_loc]
                prior_slope = np.polyfit(np.arange(len(prior_data)), prior_data['close'].values, 1)[0]
                # إعطاء مرونة طفيفة لميل الاتجاه السابق لتجنب استبعاد الفرص الحقيقية
                if prior_slope > 0.02:
                    continue
            else:
                continue

            if not (p5.price < p3.price < p1.price):
                continue
            if not (p4.price > p2.price):
                continue

            lower_slope = (p5.price - p1.price) / (p5.index - p1.index)
            lower_intercept = p1.price - lower_slope * p1.index
            
            expected_p3_price = lower_slope * p3.index + lower_intercept
            p3_error = abs(p3.price - expected_p3_price) / p3.price

            if p3_error > self.max_linearity_error:
                continue

            upper_slope = (p4.price - p2.price) / (p4.index - p2.index)
            upper_intercept = p2.price - upper_slope * p2.index

            if upper_slope <= 0 or lower_slope >= 0:
                continue

            confidence = 1.0 - (p3_error / self.max_linearity_error)

            if confidence > best_confidence:
                best_confidence = confidence
                best_pattern = PerfectBroadeningBottom(
                    start_date=p1.date,
                    entry_date=p5.date,
                    p1=p1, p2=p2, p3=p3, p4=p4, p5=p5,
                    upper_slope=upper_slope,
                    upper_intercept=upper_intercept,
                    lower_slope=lower_slope,
                    lower_intercept=lower_intercept,
                    confidence=confidence
                )

        return best_pattern


# ==============================================================================
# 3. BACKTEST ENGINE WITH REAL OPEN POSITIONS DETECTION
# ==============================================================================

def run_backtest_on_stock(
    ticker: str,
    detector: BroadeningBottomDetector,
    tv_live_dict: dict,
    years: int = 1,
    max_holding_bars: int = 40
) -> List[dict]:
    trades = []
    end_date = datetime.datetime.now()
    start_date = end_date - datetime.timedelta(days=365 * years)

    try:
        df = get_combined_stock_data(ticker, start_date, end_date, tv_live_dict)
        if df.empty or len(df) < 50:
            return trades
    except Exception:
        return trades

    latest_close_price = round(float(df['close'].iloc[-1]), 2)
    window_size = 80
    step_size = 5
    i = 0
    last_p5_date = None

    while i < len(df) - 15:
        sub_df = df.iloc[i : min(i + window_size, len(df))].copy()
        if len(sub_df) < 30:
            break

        pattern = detector.detect(sub_df)

        if pattern and (last_p5_date is None or pattern.p5.date > last_p5_date):
            last_p5_date = pattern.p5.date

            entry_date = pattern.p5.date
            entry_idx = df.index.get_loc(entry_date)
            entry_price = round(float(df['close'].loc[entry_date]), 2)
            stop_loss = round(float(pattern.p5.price * 0.98), 2)

            future_df = df.iloc[entry_idx + 1 : entry_idx + 1 + max_holding_bars]

            target_p6 = round(float(pattern.upper_value_at(pattern.p5.index + 10)), 2)

            if future_df.empty:
                trades.append({
                    'Stock Name': ticker,
                    'Entry Date': entry_date.strftime('%Y-%m-%d'),
                    'Entry Price': entry_price,
                    'Target': target_p6,
                    'Stop Loss': stop_loss,
                    'Exit Date': None,
                    'Current Price': latest_close_price,
                    'Status': 'Open'
                })
                i += step_size
                continue

            status = None
            exit_date = None

            for current_step, (date_idx, row) in enumerate(future_df.iterrows(), start=1):
                abs_idx = pattern.p5.index + current_step
                dynamic_target_p6 = pattern.upper_value_at(abs_idx)

                if row['high'] >= dynamic_target_p6:
                    status = 'Win'
                    target_p6 = round(float(dynamic_target_p6), 2)
                    exit_date = date_idx
                    break
                elif row['low'] <= stop_loss:
                    status = 'Loss'
                    exit_date = date_idx
                    break

            if status is None:
                if (entry_idx + 1 + len(future_df)) >= len(df):
                    exit_date_str = None
                    status_str = 'Open'
                else:
                    exit_date_str = future_df.index[-1].strftime('%Y-%m-%d')
                    exit_price = future_df['close'].iloc[-1]
                    status_str = 'Win' if exit_price >= entry_price else 'Loss'
            else:
                exit_date_str = exit_date.strftime('%Y-%m-%d')
                status_str = status

            trades.append({
                'Stock Name': ticker,
                'Entry Date': entry_date.strftime('%Y-%m-%d'),
                'Entry Price': entry_price,
                'Target': target_p6,
                'Stop Loss': stop_loss,
                'Exit Date': exit_date_str,
                'Current Price': latest_close_price,
                'Status': status_str
            })

            i += window_size // 2
        else:
            i += step_size

    return trades


# ==============================================================================
# 4. SINGLE-SHEET EXCEL FORMATTING
# ==============================================================================

def format_excel_file(excel_filename: str):
    wb = openpyxl.load_workbook(excel_filename)
    ws = wb.active
    ws.views.sheetView[0].showGridLines = True
    ws.views.sheetView[0].rightToLeft = False

    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    align_center = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 24
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    columns = ['Stock Name', 'Entry Date', 'Entry Price', 'Target', 'Stop Loss', 'Exit Date', 'Current Price', 'Status']

    win_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
    loss_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    open_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")

    win_font = Font(name="Calibri", size=11, bold=True, color="155724")
    loss_font = Font(name="Calibri", size=11, bold=True, color="721C24")
    open_font = Font(name="Calibri", size=11, bold=True, color="856404")

    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 20
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = align_center

            col_name = columns[col - 1] if col - 1 < len(columns) else ""

            if col_name in ['Entry Price', 'Target', 'Stop Loss', 'Current Price']:
                cell.number_format = '#,##0.00'

            if col_name == 'Status':
                val = str(cell.value).strip() if cell.value else ""
                if val == 'Win':
                    cell.fill = win_fill
                    cell.font = win_font
                elif val == 'Loss':
                    cell.fill = loss_fill
                    cell.font = loss_font
                elif val == 'Open':
                    cell.fill = open_fill
                    cell.font = open_font

    column_widths = {'A': 18, 'B': 16, 'C': 15, 'D': 15, 'E': 15, 'F': 16, 'G': 16, 'H': 14}
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    try:
        wb.save(excel_filename)
        print(f"\n📊 EXCEL EXPORT COMPLETED: {excel_filename}")
    except Exception as e:
        print(f"\n[Error saving file]: {e}")


# ==============================================================================
# 5. EXECUTION & MAIN SCANNER
# ==============================================================================

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
    "MFSC.CA", "MHOT.CA", "MILS.CA", "MIPH.CA", "MOED.CA", "MOIL.CA", "MOIN.CA", "MOSC.CA",
    "MPCI.CA", "MPCO.CA", "MPRC.CA", "MTIE.CA", "NAHO.CA", "NARE.CA", "NCCW.CA", "NCGC.CA",
    "NEDA.CA", "NHPS.CA", "NINH.CA", "NIPH.CA", "OBRI.CA", "OCAP.CA", "OCDI.CA", "OCPH.CA",
    "ODIN.CA", "OFH.CA", "OIH.CA", "OLFI.CA", "ORAS.CA", "ORHD.CA", "ORWE.CA", "PHAR.CA",
    "PHDC.CA", "PHGC.CA", "PHTV.CA", "POUL.CA", "PRCL.CA", "PRDC.CA", "PRMH.CA", "QNBE.CA",
    "RACC.CA", "RAKT.CA", "RAYA.CA", "RKAZ.CA", "RMDA.CA", "RMTV.CA", "ROTO.CA", "RREI.CA",
    "RTVC.CA", "RUBX.CA", "SAUD.CA", "SCEM.CA", "SCFM.CA", "SCTS.CA", "SDTI.CA", "SEIG.CA",
    "SIEG.CA", "SIPC.CA", "SKPC.CA", "SMFR.CA", "SNFC.CA", "SPIN.CA", "SPMD.CA", "SUCE.CA",
    "SUGR.CA", "SVCE.CA", "SWDY.CA", "TALM.CA", "TANM.CA", "TAQA.CA", "TMGH.CA", "TORA.CA",
    "TWSA.CA", "TYCN.CA", "UBEE.CA", "UEFM.CA", "UEGC.CA", "UNIP.CA", "UNIT.CA", "UPMS.CA",
    "UTOP.CA", "VALU.CA", "VERT.CA", "VLMR.CA", "VLMRA.CA", "WCDF.CA", "WKOL.CA", "ZEOT.CA",
    "ZMID.CA"
]

if __name__ == "__main__":
    detector = BroadeningBottomDetector(order=4, max_linearity_error=0.020)
    all_trades = []

    print("🔍 Starting Standalone Scan: Broadening Bottoms Pattern...")
    print("Fetching TradingView live prices for EGX stocks...")
    tv_live_dict = fetch_tradingview_live_data(egyptian_stocks)

    print(f"Scanning {len(egyptian_stocks)} stocks...")
    for idx, ticker in enumerate(egyptian_stocks, 1):
        print(f"[{idx}/{len(egyptian_stocks)}] Processing: {ticker}...")
        trades = run_backtest_on_stock(ticker, detector, tv_live_dict, years=1)
        if trades:
            print(f"  └─ 🟢 Found {len(trades)} trade(s) for {ticker}")
        all_trades.extend(trades)
        
        # مهلة قصيرة لحماية السيرفر من الحظر التلقائي أثناء التشغيل السحابي
        time.sleep(0.2)

    df_trades = pd.DataFrame(all_trades)

    if not df_trades.empty:
        df_trades.drop_duplicates(subset=['Stock Name', 'Entry Date'], inplace=True)
        df_trades['Temp Date'] = pd.to_datetime(df_trades['Entry Date'])
        df_trades.sort_values(by='Temp Date', ascending=True, inplace=True)
        df_trades.drop(columns=['Temp Date'], inplace=True)

        excel_filename = "Broadening_Bottoms_Scan_Results.xlsx"
        
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            df_trades.to_excel(writer, sheet_name='Trades List', index=False)

        format_excel_file(excel_filename)
        
        print("\n" + "=" * 80)
        print("🎯 BROADENING BOTTOMS SCAN RESULTS")
        print("=" * 80)
        
        open_trades = df_trades[df_trades['Status'] == 'Open']
        if not open_trades.empty:
            print(f"\n[!] Active / Open Positions Found ({len(open_trades)}):")
            print(open_trades[['Stock Name', 'Entry Date', 'Entry Price', 'Target', 'Stop Loss', 'Current Price', 'Status']].to_string(index=False))
        else:
            print("\n[!] No active 'Open' positions found right now. Showing recent historical trades:")
            print(df_trades.tail(10)[['Stock Name', 'Entry Date', 'Entry Price', 'Target', 'Stop Loss', 'Exit Date', 'Status']].to_string(index=False))
            
        print("\n" + "=" * 80)
    else:
        print("\n❌ No Broadening Bottom patterns detected in the current period.")
