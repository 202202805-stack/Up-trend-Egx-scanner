# =====================================================================
# EGX TRIPLE BOTTOM PATTERN BACKTESTER (BULKOWSKI ENHANCED + METRICS)
# UPDATED FOR GITHUB ACTIONS & CI/CD PIPELINES
# =====================================================================

from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
import datetime
import os
import sys
import warnings

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import requests
from scipy.signal import argrelextrema
from tqdm import tqdm
import yfinance as yf

warnings.filterwarnings("ignore")


@dataclass
class TBConfig:
  pivot_order: int = 5
  min_separation_bars: int = 10
  price_tolerance_pct: float = 3.0
  max_width_bars: int = 78
  min_rally_pct: float = 3.0
  require_prior_downtrend: bool = True
  downtrend_lookback: int = 60
  downtrend_drop_pct: float = 15.0
  breakout_lookforward: int = 90
  max_holding_bars: int = 120  # أقصى مدة احتفاظ بالشموع
  sma_regime_len: int = 50
  stop_buffer_pct: float = 0.005


EGX_TICKERS = [
    "AALR.CA",
    "ABUK.CA",
    "ACAMD.CA",
    "ACAP.CA",
    "ACGC.CA",
    "ACTF.CA",
    "ADCI.CA",
    "ADIB.CA",
    "ADPC.CA",
    "ADRI.CA",
    "AFDI.CA",
    "AFMC.CA",
    "AIDC.CA",
    "AIFI.CA",
    "AIH.CA",
    "AJWA.CA",
    "ALCN.CA",
    "ALEX.CA",
    "ALUM.CA",
    "AMER.CA",
    "AMES.CA",
    "AMIA.CA",
    "AMII.CA",
    "AMOC.CA",
    "AMPI.CA",
    "APSW.CA",
    "ARAB.CA",
    "ARCC.CA",
    "AREH.CA",
    "ASCM.CA",
    "ASPI.CA",
    "ATLC.CA",
    "ATQA.CA",
    "AXPH.CA",
    "BIDI.CA",
    "BIGP.CA",
    "BINV.CA",
    "BIOC.CA",
    "BONY.CA",
    "BTFH.CA",
    "CAED.CA",
    "CANA.CA",
    "CCAP.CA",
    "CCRS.CA",
    "CEFM.CA",
    "CERA.CA",
    "CFGH.CA",
    "CICH.CA",
    "CIEB.CA",
    "CIRA.CA",
    "CLHO.CA",
    "CNFN.CA",
    "COMI.CA",
    "COPR.CA",
    "COSG.CA",
    "CPCI.CA",
    "CPME.CA",
    "CRST.CA",
    "CSAG.CA",
    "DAPH.CA",
    "DCRC.CA",
    "DEIN.CA",
    "DGTZ.CA",
    "DOMT.CA",
    "DSCW.CA",
    "DTPP.CA",
    "EALR.CA",
    "EASB.CA",
    "EAST.CA",
    "EBSC.CA",
    "ECAP.CA",
    "EDFM.CA",
    "EEII.CA",
    "EFIC.CA",
    "EFID.CA",
    "EFIH.CA",
    "EGAL.CA",
    "EGAS.CA",
    "EGBE.CA",
    "EGCH.CA",
    "EGREF.CA",
    "EGSA.CA",
    "EGTS.CA",
    "EHDR.CA",
    "ELAB.CA",
    "ELEC.CA",
    "ELKA.CA",
    "ELNA.CA",
    "ELSH.CA",
    "ELWA.CA",
    "EMFD.CA",
    "ENGC.CA",
    "EOSB.CA",
    "EPCO.CA",
    "EPPK.CA",
    "ETEL.CA",
    "ETRS.CA",
    "EXPA.CA",
    "FAIT.CA",
    "FAITA.CA",
    "FCMD.CA",
    "FIRE.CA",
    "FNAR.CA",
    "FTNS.CA",
    "FWRY.CA",
    "GBCO.CA",
    "GDWA.CA",
    "GGCC.CA",
    "GGRN.CA",
    "GIHD.CA",
    "GMCI.CA",
    "GOUR.CA",
    "GPIM.CA",
    "GRCA.CA",
    "GSSC.CA",
    "GTEX.CA",
    "GTHE.CA",
    "GTWL.CA",
    "HBCO.CA",
    "HDBK.CA",
    "HELI.CA",
    "HRHO.CA",
    "IBCT.CA",
    "ICFC.CA",
    "ICID.CA",
    "IDRE.CA",
    "IEEC.CA",
    "IFAP.CA",
    "INEG.CA",
    "INFI.CA",
    "IRON.CA",
    "ISMA.CA",
    "ISMQ.CA",
    "ISPH.CA",
    "JUFO.CA",
    "KABO.CA",
    "KORA.CA",
    "KRDI.CA",
    "KWIN.CA",
    "KZPC.CA",
    "LCSW.CA",
    "LKGP.CA",
    "LUTS.CA",
    "MAAL.CA",
    "MASR.CA",
    "MBEG.CA",
    "MBSC.CA",
    "MCQE.CA",
    "MCRO.CA",
    "MENA.CA",
    "MEPA.CA",
    "MFPC.CA",
    "MFSC.CA",
    "MHOT.CA",
    "MICH.CA",
    "MILS.CA",
    "MIPH.CA",
    "MOED.CA",
    "MOIL.CA",
    "MOIN.CA",
    "MOSC.CA",
    "MPCI.CA",
    "MPCO.CA",
    "MPRC.CA",
    "MTIE.CA",
    "NAHO.CA",
    "NARE.CA",
    "NCCW.CA",
    "NCGC.CA",
    "NEDA.CA",
    "NHPS.CA",
    "NINH.CA",
    "NIPH.CA",
    "OBRI.CA",
    "OCAP.CA",
    "OCDI.CA",
    "OCPH.CA",
    "ODIN.CA",
    "OFH.CA",
    "OIH.CA",
    "OLFI.CA",
    "ORAS.CA",
    "ORHD.CA",
    "ORWE.CA",
    "PHAR.CA",
    "PHDC.CA",
    "PHGC.CA",
    "PHTV.CA",
    "POUL.CA",
    "PRCL.CA",
    "PRDC.CA",
    "PRMH.CA",
    "QNBE.CA",
    "RACC.CA",
    "RAKT.CA",
    "RAYA.CA",
    "RKAZ.CA",
    "RMDA.CA",
    "RMTV.CA",
    "ROTO.CA",
    "RREI.CA",
    "RTVC.CA",
    "RUBX.CA",
    "SAUD.CA",
    "SCEM.CA",
    "SCFM.CA",
    "SCTS.CA",
    "SDTI.CA",
    "SEIG.CA",
    "SIEG.CA",
    "SIPC.CA",
    "SKPC.CA",
    "SMFR.CA",
    "SNFC.CA",
    "SPIN.CA",
    "SPMD.CA",
    "SUCE.CA",
    "SUGR.CA",
    "SVCE.CA",
    "SWDY.CA",
    "TALM.CA",
    "TANM.CA",
    "TAQA.CA",
    "TMGH.CA",
    "TORA.CA",
    "TWSA.CA",
    "TYCN.CA",
    "UBEE.CA",
    "UEFM.CA",
    "UEGC.CA",
    "UNIP.CA",
    "UNIT.CA",
    "UPMS.CA",
    "UTOP.CA",
    "VALU.CA",
    "VERT.CA",
    "VLMR.CA",
    "VLMRA.CA",
    "WCDF.CA",
    "WKOL.CA",
    "ZEOT.CA",
    "ZMID.CA",
]


def fetch_tradingview_live_data():
  url = "https://scanner.tradingview.com/egypt/scan"
  payload = {
      "filter": [{"left": "name", "operation": "nempty"}],
      "options": {"active_symbols_only": True},
      "columns": ["name", "open", "high", "low", "close", "volume"],
      "sort": {"sortBy": "name", "sortOrder": "asc"},
      "range": [0, 300],
  }
  headers = {"User-Agent": "Mozilla/5.0"}
  tv_dict = {}

  try:
    res = requests.post(url, json=payload, headers=headers, timeout=10)
    data = res.json().get("data", [])
    today_dt = pd.Timestamp(datetime.date.today())

    for item in data:
      sym = item["s"].replace("EGX:", "")
      d = item["d"]
      open_p = float(d[1]) if d[1] is not None else 0.0
      high_p = float(d[2]) if d[2] is not None else open_p
      low_p = float(d[3]) if d[3] is not None else open_p
      close_p = float(d[4]) if d[4] is not None else open_p
      vol_p = float(d[5]) if d[5] is not None else 0.0

      if close_p > 0:
        tv_dict[f"{sym}.CA"] = {
            "Date": today_dt,
            "Open": round(open_p, 3),
            "High": round(high_p, 3),
            "Low": round(low_p, 3),
            "Close": round(close_p, 3),
            "Volume": vol_p,
        }
  except Exception as e:
    print(f"⚠️ Warning: Could not fetch TradingView live data: {e}")

  return tv_dict


def merge_tv_live_data(df: pd.DataFrame, tv_row: dict) -> pd.DataFrame:
  if not tv_row:
    return df

  tv_date = tv_row["Date"]
  last_df_date = df.index[-1]

  if tv_date.date() > last_df_date.date():
    new_row = pd.DataFrame(
        [{
            "Open": tv_row["Open"],
            "High": tv_row["High"],
            "Low": tv_row["Low"],
            "Close": tv_row["Close"],
            "Volume": tv_row["Volume"],
        }],
        index=[tv_date],
    )
    df = pd.concat([df, new_row])
  elif tv_date.date() == last_df_date.date():
    df.loc[last_df_date, ["Open", "High", "Low", "Close", "Volume"]] = [
        tv_row["Open"],
        tv_row["High"],
        tv_row["Low"],
        tv_row["Close"],
        tv_row["Volume"],
    ]

  return df


def get_swing_indices(series: pd.Series, order: int, kind: str) -> np.ndarray:
  f = np.less if kind == "low" else np.greater
  return argrelextrema(series.values, f, order=order)[0]


def enforce_separation(idx: np.ndarray, min_sep: int) -> np.ndarray:
  kept, last = [], -(10**9)
  for i in idx:
    if i - last >= min_sep:
      kept.append(i)
      last = i
  return np.array(kept, dtype=int)


def fetch_single_ticker(
    ticker: str, start_date: str, end_date: str, tv_dict: dict
):
  try:
    t = yf.Ticker(ticker)
    df = t.history(start=start_date, end=end_date, auto_adjust=False)

    if ticker in tv_dict:
      df = merge_tv_live_data(df, tv_dict[ticker])

    if df.empty or len(df) < 60:
      return ticker, None

    if isinstance(df.columns, pd.MultiIndex):
      df.columns = df.columns.get_level_values(0)

    df = df[["Open", "High", "Low", "Close", "Volume"]].dropna()
    df.columns = [col.capitalize() for col in df.columns]

    if (df["High"] == df["Low"]).all():
      return ticker, None

    return ticker, df
  except Exception:
    return ticker, None


def download_all_data(
    tickers: list, years: int = 1, max_workers: int = 10
) -> dict:
  end_dt = datetime.date.today()
  start_dt = end_dt - datetime.timedelta(days=years * 365 + 60)

  print("1. Fetching live session data from TradingView...")
  tv_dict = fetch_tradingview_live_data()

  data_store = {}
  print(f"2. Downloading historical stock data for {len(tickers)} assets...")

  with ThreadPoolExecutor(max_workers=max_workers) as executor:
    futures = {
        executor.submit(
            fetch_single_ticker, t, str(start_dt), str(end_dt), tv_dict
        ): t
        for t in tickers
    }
    for future in tqdm(
        as_completed(futures), total=len(tickers), desc="Downloading Stock Data"
    ):
      ticker, df = future.result()
      if df is not None:
        data_store[ticker] = df

  print(f"✅ Successfully processed {len(data_store)} stocks.")
  return data_store


def analyze_stock_trades(ticker: str, df: pd.DataFrame, cfg: TBConfig) -> list:
  n = len(df)
  if n < cfg.sma_regime_len + 20:
    return []

  sma200 = df["Close"].rolling(cfg.sma_regime_len).mean()

  sw_low = enforce_separation(
      get_swing_indices(df["Low"], cfg.pivot_order, "low"),
      cfg.min_separation_bars,
  )

  highs, lows, closes, opens = (
      df["High"].values,
      df["Low"].values,
      df["Close"].values,
      df["Open"].values,
  )
  vols = df["Volume"].values
  dates = df.index

  detected_patterns = []
  m = len(sw_low)

  for a in range(m):
    i = sw_low[a]

    if cfg.require_prior_downtrend:
      lookback_start = max(0, i - cfg.downtrend_lookback)
      pre_high = highs[lookback_start:i].max()
      pre_drop = (pre_high - lows[i]) / pre_high * 100.0
      if pre_drop < cfg.downtrend_drop_pct:
        continue

    for b in range(a + 1, m):
      j = sw_low[b]
      if j - i < cfg.min_separation_bars:
        continue

      for cix in range(b + 1, m):
        k = sw_low[cix]
        if k - j < cfg.min_separation_bars:
          continue
        if k - i > cfg.max_width_bars:
          break

        l1, l2, l3 = lows[i], lows[j], lows[k]
        avg_bottom = (l1 + l2 + l3) / 3.0

        if max(
            abs(l1 - avg_bottom), abs(l2 - avg_bottom), abs(l3 - avg_bottom)
        ) / avg_bottom > (cfg.price_tolerance_pct / 100.0):
          continue

        if l3 < l2:
          continue

        v2, v3 = vols[j], vols[k]
        if v3 >= v2:
          continue

        h1 = highs[i + 1 + np.argmax(highs[i + 1 : j + 1])]
        h2 = highs[j + 1 + np.argmax(highs[j + 1 : k + 1])]
        neckline = max(h1, h2)

        bo = None
        for t in range(k + 1, min(k + 1 + cfg.breakout_lookforward, n)):
          if closes[t] > neckline:
            bo = t
            break
        if bo is None or bo + 1 >= n:
          continue

        s200 = sma200.values[bo]
        if np.isnan(s200) or closes[bo] < s200:
          continue

        min_low = min(l1, l2, l3)
        pattern_height = neckline - min_low
        target_price = neckline + pattern_height
        stop_loss = min_low * (1 - cfg.stop_buffer_pct)

        detected_patterns.append({
            "ticker": ticker,
            "b1_idx": i,
            "b2_idx": j,
            "b3_idx": k,
            "bo_idx": bo,
            "b1_date": dates[i],
            "b2_date": dates[j],
            "b3_date": dates[k],
            "breakout_date": dates[bo],
            "neckline": neckline,
            "min_low": min_low,
            "stop_price": stop_loss,
            "target_price": target_price,
            "pattern_height": pattern_height,
        })

  if not detected_patterns:
    return []

  accepted = []
  for p in detected_patterns:
    if all(
        abs(p["bo_idx"] - a["bo_idx"]) >= cfg.min_separation_bars
        for a in accepted
    ):
      accepted.append(p)

  trades = []
  current_market_price = round(float(closes[-1]), 3)

  for p in accepted:
    bo_idx = p["bo_idx"]
    entry_idx = bo_idx + 1
    entry_date = dates[entry_idx]
    entry_price = round(float(opens[entry_idx]), 3)

    stop = round(float(p["stop_price"]), 3)
    target = round(float(p["target_price"]), 3)

    exit_date = None
    exit_price = None

    # تحديد أقصى مؤشر شمعة مسموح به للاحتفاظ
    max_holding_idx = min(n - 1, entry_idx + cfg.max_holding_bars)

    for curr_idx in range(entry_idx, max_holding_idx + 1):
      curr_low = lows[curr_idx]
      curr_high = highs[curr_idx]

      # 1. فحص الوصول للهدف
      if curr_high >= target:
        exit_date = dates[curr_idx].strftime("%Y-%m-%d")
        exit_price = target if opens[curr_idx] < target else opens[curr_idx]
        break
      # 2. فحص كسر وقف الخسارة
      elif curr_low <= stop:
        exit_date = dates[curr_idx].strftime("%Y-%m-%d")
        exit_price = stop if opens[curr_idx] > stop else opens[curr_idx]
        break

    # 3. إغلاق الصفقة بالزمن عند انتهاء أقصى مدة احتفاظ
    if exit_date is None and max_holding_idx < n - 1:
      exit_date = dates[max_holding_idx].strftime("%Y-%m-%d")
      exit_price = round(float(closes[max_holding_idx]), 3)

    # تسجيل الحالة بالإنجليزية والربح/الخسارة
    if exit_date is None:
      exit_date_str = ""
      pnl_ratio = (current_market_price - entry_price) / entry_price
      status = "OPEN"
    else:
      exit_date_str = exit_date
      pnl_ratio = (exit_price - entry_price) / entry_price
      status = "WIN" if pnl_ratio > 0 else "LOSS"

    trades.append({
        "Stock Name": p["ticker"],
        "Entry Date": entry_date.strftime("%Y-%m-%d"),
        "Entry Price": entry_price,
        "Target": target,
        "Stop Loss": stop,
        "Exit Date": exit_date_str,
        "Current Price": current_market_price,
        "PnP_Ratio": pnl_ratio,
        "Status": status,
    })

  return trades


def run_full_backtest():
  cfg = TBConfig()
  data_store = download_all_data(EGX_TICKERS, years=1)

  all_trades = []
  print("\n🔍 Scanning charts for Triple Bottom patterns...")

  for ticker in tqdm(data_store.keys(), desc="Analyzing Charts"):
    df = data_store[ticker]
    trades = analyze_stock_trades(ticker, df, cfg)
    all_trades.extend(trades)

  if not all_trades:
    print(
        "❌ No confirmed high-quality patterns found matching Bulkowski"
        " filters."
    )
    return

  trades_df = pd.DataFrame(all_trades)
  trades_df.drop_duplicates(
      subset=["Stock Name", "Entry Date"], keep="first", inplace=True
  )
  trades_df.sort_values(by="Entry Date", ascending=True, inplace=True)
  trades_df.reset_index(drop=True, inplace=True)

  total_trades = len(trades_df)
  winning_df = trades_df[trades_df["Status"] == "WIN"]
  losing_df = trades_df[trades_df["Status"] == "LOSS"]
  open_df = trades_df[trades_df["Status"] == "OPEN"]

  winning_trades = len(winning_df)
  losing_trades = len(losing_df)
  open_trades = len(open_df)

  closed_trades = winning_trades + losing_trades
  win_rate = (winning_trades / closed_trades * 100) if closed_trades > 0 else 0

  avg_win = winning_df["PnP_Ratio"].mean() * 100 if winning_trades > 0 else 0.0
  avg_loss = losing_df["PnP_Ratio"].mean() * 100 if losing_trades > 0 else 0.0

  avg_pnl = trades_df["PnP_Ratio"].mean() * 100
  total_cum_pnl = trades_df["PnP_Ratio"].sum() * 100

  # الأعمدة الـ 8 المحددة بالإنجليزية
  output_columns = [
      "Stock Name",
      "Entry Date",
      "Entry Price",
      "Target",
      "Stop Loss",
      "Exit Date",
      "Current Price",
      "Status",
  ]
  excel_df = trades_df[output_columns]

  output_filename = "EGX_Triple_Bottom_Bulkowski_Results.xlsx"
  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  header_fill = PatternFill(
      start_color="1B365D", end_color="1B365D", fill_type="solid"
  )
  header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  data_font = Font(name="Calibri", size=11, bold=False, color="000000")

  center_align = Alignment(horizontal="center", vertical="center")
  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )

  ws_trades = wb.create_sheet(title="All Trades List")
  ws_trades.views.sheetView[0].showGridLines = True

  for r in dataframe_to_rows(excel_df, index=False, header=True):
    ws_trades.append(r)

  for cell in ws_trades[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_align

  # تنسيق ورقة الصفقات ووضع الألوان المناسبة لكل حالة
  for row_idx in range(2, ws_trades.max_row + 1):
    status_val = ws_trades.cell(row=row_idx, column=8).value

    for col_idx in range(1, 9):
      cell = ws_trades.cell(row=row_idx, column=col_idx)
      cell.alignment = center_align
      cell.font = data_font
      cell.border = thin_border

      # تنسيق الأرقام والأسعار
      if col_idx in [3, 4, 5, 7] and isinstance(cell.value, (int, float)):
        cell.number_format = "0.000"

      # تنسيق خلية الحالة Status بالألوان
      if col_idx == 8:
        if status_val == "WIN":
          cell.fill = PatternFill(start_color="DCFCE7", fill_type="solid")
          cell.font = Font(
              name="Calibri", size=11, bold=True, color="166534"
          )
        elif status_val == "LOSS":
          cell.fill = PatternFill(start_color="FEE2E2", fill_type="solid")
          cell.font = Font(
              name="Calibri", size=11, bold=True, color="991B1B"
          )
        elif status_val == "OPEN":
          cell.fill = PatternFill(start_color="FEF3C7", fill_type="solid")
          cell.font = Font(
              name="Calibri", size=11, bold=True, color="92400E"
          )

  # تعيين عرض الأعمدة الـ 8 تلقائياً
  col_widths = {
      "A": 16,
      "B": 16,
      "C": 14,
      "D": 14,
      "E": 14,
      "F": 16,
      "G": 16,
      "H": 14,
  }
  for col_letter, width in col_widths.items():
    ws_trades.column_dimensions[col_letter].width = width

  summary_df = pd.DataFrame({
      "Metric": [
          "Backtest Period",
          "Total Assets Scanned",
          "Total Trades Detected",
          "Open Positions Currently Active",
          "Winning Trades (Closed)",
          "Losing Trades (Closed)",
          "Closed Trades Win Rate (%)",
          "Average Win (%)",
          "Average Loss (%)",
          "Average Return per Trade (%)",
          "Cumulative Return (Sum PnL %)",
      ],
      "Value": [
          "1 Year (Live TV Integrated)",
          len(data_store),
          total_trades,
          open_trades,
          winning_trades,
          losing_trades,
          f"{win_rate:.2f}%",
          f"{avg_win:.2f}%",
          f"{avg_loss:.2f}%",
          f"{avg_pnl:.2f}%",
          f"{total_cum_pnl:.2f}%",
      ],
  })

  ws_summary = wb.create_sheet(title="Backtest Summary")
  ws_summary.views.sheetView[0].showGridLines = True

  for r in dataframe_to_rows(summary_df, index=False, header=True):
    ws_summary.append(r)

  for cell in ws_summary[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_align

  for row in ws_summary.iter_rows(
      min_row=2, max_row=ws_summary.max_row, min_col=1, max_col=2
  ):
    for cell in row:
      cell.font = data_font
      cell.alignment = center_align
      cell.border = thin_border

  for col in ws_summary.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_summary.column_dimensions[col_letter].width = max(max_len + 5, 20)

  wb.save(output_filename)

  print("\n" + "=" * 60)
  print("🎉 BULKOWSKI STRATEGY SCAN & BACKTEST COMPLETED!")
  print(f"📊 Total Trades Found (1 Year): {total_trades}")
  print(f"🔵 Active Open Positions:       {open_trades}")
  print(f"🏆 Closed Trades Win Rate:     {win_rate:.2f}%")
  print(f"🟢 Average Win PnL:             {avg_win:.2f}%")
  print(f"🔴 Average Loss PnL:            {avg_loss:.2f}%")
  print(f"📈 Average Return/Trade:        {avg_pnl:.2f}%")
  print(f"📁 Excel Saved As:              {output_filename}")
  print("=" * 60)


if __name__ == "__main__":
  run_full_backtest()
