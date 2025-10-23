#Required for this module 
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl  # Excel engine for .xlsx
from pathlib import Path
from typing import Union, IO, Optional


##plt.style.use("default")

import matplotlib as mpl
mpl.rcParams["font.family"] = "serif"
mpl.rc("xtick", labelsize=12)
mpl.rc("ytick", labelsize=12)

plt.style.use("seaborn-v0_8-white")



# Optional / nice-to-have
try:
    import seaborn as sns
except Exception:
    sns = None

# Data & utils (only used in other notebooks/features)
try:
    import yfinance as yf
    import datetime as dt
except Exception:
    yf = None
    dt = None

try:
    import scipy as spy
    from scipy.optimize import minimize
except Exception:
    spy = None
    minimize = None

try:
    from sklearn.metrics import calinski_harabasz_score, silhouette_score, davies_bouldin_score
    from sklearn.cluster import KMeans
except Exception:
    calinski_harabasz_score = silhouette_score = davies_bouldin_score = KMeans = None

try:
    from pylatex import Document, Section, Subsection, Tabular, Math, TikZ, Axis, Plot, Figure, Matrix, Alignat
    from pylatex.utils import italic
except Exception:
    Document = Section = Subsection = Tabular = Math = TikZ = Axis = Plot = Figure = Matrix = Alignat = italic = None

try:
    from plottable import Table
except Exception:
    Table = None

try:
    import jupyter_capture_output
except Exception:
    jupyter_capture_output = None

try:
    from openpyxl import Workbook  
except Exception:
    Workbook = None

import math

ExcelLike = Union[str, Path, IO[bytes]]

def _excel_has_sheet(x: ExcelLike, sheet: str) -> bool:
    try:
        return sheet in pd.ExcelFile(x).sheet_names
    except Exception:
        return False

def hello():

    x = int(input(f"Which analysis you\'re looking to do today:\n\
    1. Country.\n\
    2. Sector.\n\
    ** Enter the Number from 1 or 2 **"))
    
    countryname = input(' Please name the country which you are focusing today').upper() 
            
    return x, countryname

def master_data(
    prices_file: Optional[ExcelLike] = None,
    meta_file:   Optional[ExcelLike] = None,
    prices_sheet: str = "SP",
    meta_sheet:   Optional[str] = None,
):
    """
    Loads price & metadata either from uploaded files (paths or BytesIO)
    OR from your conventional filenames:
      - "<countryname>_Share Price_Combined.xlsx" (sheet 'SP')
      - "1. Tic_Global.xlsx" (sheet '<countryname>')
    Sets module globals rdata/df0 so other functions work unchanged.
    Returns: (rdata, df0)
    """
    global rdata, df0, countryname

    # Resolve sheet name for meta
    if meta_sheet is None:
        if countryname is None:
            raise ValueError("countryname not set. Call hello() or set masketer.core.countryname.")
        meta_sheet = countryname

    # Resolve files (uploaded or defaults)
    if prices_file is None:
        if countryname is None:
            raise ValueError("countryname not set. Call hello() or set masketer.core.countryname.")
        prices_file = f"{countryname}_Share Price_Combined.xlsx"
    if meta_file is None:
        meta_file = "1. Tic_Global.xlsx"

    # Sheet checks (replaces the invalid '...' line)
    if not _excel_has_sheet(prices_file, prices_sheet):
        raise KeyError(f"'{prices_sheet}' sheet missing in prices workbook.")
    if not _excel_has_sheet(meta_file, meta_sheet):
        raise KeyError(f"'{meta_sheet}' sheet missing in meta workbook.")

    # ---- read prices & normalize date
    rdata = pd.read_excel(pd.ExcelFile(prices_file), prices_sheet)
    if "Date" in rdata.columns:
        rdata["Date"] = pd.to_datetime(rdata["Date"], errors="coerce")
        rdata = rdata.sort_values("Date").set_index("Date")
    else:
        # if Date was already index or different label
        rdata.index = pd.to_datetime(rdata.index, errors="coerce")
        rdata = rdata.sort_index()

    # Keep your original reversal
    rdata = rdata.reindex(index=rdata.index[::-1])

    # read metadata
    df0 = pd.read_excel(pd.ExcelFile(meta_file), meta_sheet)
    if "Company2" in df0.columns:
        df0 = df0.drop(columns=["Company2"])

    # create placeholders
    _cn = (countryname or meta_sheet or "OUTPUT")
    pd.DataFrame().to_excel(f"{_cn}_1-Return analysis.xlsx")
    pd.DataFrame().to_excel(f"{_cn}_2-Summary analysis.xlsx")

    return rdata, df0
# Globals variables
countryname = globals().get("countryname", None)
rdata = globals().get("rdata", None)
df0 = globals().get("df0", None)
x = globals().get("x", None)

# Make master_data set the globals so other functions see them
def _wrap_master_data(_orig):
    def _inner(*args, **kwargs):
        r, d = _orig(*args, **kwargs)
        globals()["rdata"] = r
        globals()["df0"] = d
        return r, d
    return _inner

# replace master_data with wrapped version
master_data = _wrap_master_data(master_data)

#  Analysis functions 

def country_analysis(countryname):
    global rdata
    if rdata is None:
        raise RuntimeError("rdata not loaded. Call master_data() first.")
    R = rdata.pct_change(fill_method=None)
    R.to_excel(f"{countryname}_return analysis.xlsx")

    S = pd.DataFrame(R.describe(percentiles=[0.01, 0.1, 0.5, 0.9, 0.99]))
    S.to_excel(f"{countryname}_summary analysis.xlsx")
    return R, S


def sector_analysis(countryname):   # exisitng dataset for the entire country is available
    global rdata, df0
    if rdata is None or df0 is None:
        raise RuntimeError("rdata/df0 not loaded. Call master_data() first.")

    sector = input('Please specify the sector you wish to focus today ? ').strip()
    print(f"Tickers of Sector: {sector}")
    ticker = list(df0.loc[df0['Sector'] == sector, 'Ticker'])

    # --- STEP 1 RETURN ---
    cols = [t for t in ticker if t in rdata.columns]
    if not cols:
        raise ValueError(f"No overlapping tickers for sector '{sector}' in rdata.")
    R = rdata[cols].pct_change(fill_method=None)

    # --- STEP 3 SUMMARY ---
    S = pd.DataFrame(R.describe(percentiles=[0.01, 0.1, 0.5, 0.9, 0.99]))
    return sector, ticker, R, S


def KPIs(R, S, sector, ticker):  # keep your plotting styles
    a = (((R + 1).cumprod()) - 1).dropna().iloc[-1:].round(3)
    a.T.plot(kind='bar', color='k', figsize=(12, 4),
             title='Distribution of Cumulative Returns',
             fontsize=10, grid=True, legend=False)
    plt.savefig(f"cum_ret_dis_{sector}_{countryname}", bbox_inches='tight')

    plt.figure(figsize=(12, 15))
    plt.suptitle("Stocks' KPIS for Investors", fontsize=16, color='r', y=0.93)

    plt.subplot(221)
    plt.title('Average Return on Stocks (Daily)', fontsize=14, y=1.01)
    S.loc['mean'].sort_values(ascending=False).plot(
        kind='barh', fontsize=10,
        color=['r', 'g', 'b', 'c', 'k'], edgecolor='k', linestyle='--', grid=True
    )

    plt.subplot(222)
    plt.title('Average Voltality in Stock Returns (Daily)', fontsize=14, y=1.01)
    S.loc['std'].sort_values(ascending=False).plot(
        kind='barh', fontsize=10,
        color=['r', 'g', 'b', 'c', 'k'], edgecolor='k', linestyle='--', grid=True
    )

    plt.savefig(f"risk_ret_{sector}_{countryname}", bbox_inches='tight')
    plt.show()

    want = ['1%', '10%', '50%', '90%', '99%']
    idx = [i for i in S.index.astype(str) if i in want]
    if idx:
        S.loc[idx].T.plot(kind='box', vert=True, figsize=(12, 4),
                          title='Distribution of Daily Returns', fontsize=10,
                          grid=True, color='k', style='-b', widths=0.1,
                          showcaps=True, showbox=True, showmeans=True,
                          boxprops=dict(linestyle='-', linewidth=1.5))
    else:
        R.quantile([0.01, 0.1, 0.5, 0.9, 0.99]).T.plot(
            kind='box', vert=True, figsize=(12, 4),
            title='Distribution of Daily Returns', fontsize=10, grid=True
        )
    plt.savefig(f"high_low_ret_{sector}_{countryname}", bbox_inches='tight')
    return
def _to_numeric_clean(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce to numeric, drop all-NaN and constant columns."""
    x = df.apply(pd.to_numeric, errors="coerce")
    x = x.dropna(axis=1, how="all")           # remove columns that are entirely NaN
    x = x.loc[:, x.nunique(dropna=True) > 1]  # remove constant columns (no variance)
    return x

def corr_ana(R: pd.DataFrame, sector: str, rdata: pd.DataFrame):
    """Correlation heatmaps for prices and returns (auto-cleans numeric columns)."""
    if sns is None:
        raise ImportError("seaborn is required for corr_ana(); pip install seaborn or masketer[notebook].")

    #clean data like you did in Jupyter
    rdata_clean  = _to_numeric_clean(rdata)
    R_clean      = _to_numeric_clean(R)

    # align columns (R columns must exist in rdata for price corr)
    cols = [c for c in R_clean.columns if c in rdata_clean.columns]
    if not cols:
        raise ValueError("No overlapping numeric tickers between rdata and R after cleaning.")
    rdata_clean = rdata_clean[cols]
    R_clean     = R_clean[cols]

    plt.figure(figsize=(15, 12))

    # Price correlation
    plt.subplot(211)
    plt.title("Correlation of Stock Prices", fontsize=12)
    corr_p = rdata_clean.corr()
    mask = np.triu(np.ones_like(corr_p, dtype=bool))
    cmap = sns.diverging_palette(230, 20, s=75, l=50, sep=2, n=5, center="light", as_cmap=True)
    sns.heatmap(corr_p, mask=mask, cmap=cmap, vmax=2, center=0, vmin=-2,
                square=False, linewidths=1, cbar_kws={"shrink": 1}, annot=False)

    # Return correlation
    plt.subplot(212)
    plt.title("Correlation of Stock Returns", fontsize=12)
    corr_r = R_clean.corr()
    mask = np.triu(np.ones_like(corr_r, dtype=bool))
    sns.heatmap(corr_r, mask=mask, cmap=cmap, vmax=2, center=0, vmin=-2,
                square=False, linewidths=1, cbar_kws={"shrink": 1}, annot=False)

    plt.savefig(f"corr_matrix_{sector}_{countryname}", bbox_inches="tight")
    plt.show()
    return


# Orchestrators

def data():
    """Run hello() then load datasets. Returns (rdata, df0)."""
    global x, countryname, rdata, df0
    x, countryname = hello()
    rdata, df0 = master_data()     # relies on global countryname
    return rdata, df0

def analysis():
    """Branch on x; returns (R, S, ticker, sector)."""
    if x == 1:
        R, S = country_analysis(countryname)
        ticker, sector = " ", " "
    elif x == 2:
        sector, ticker, R, S = sector_analysis(countryname)
        print(ticker)
    else:
        raise ValueError("x must be 1 (Country) or 2 (Sector).")
    return R, S, ticker, sector

def charts():
    """Produce KPI & correlation charts; returns (x, R, S)."""
    R, S, ticker, sector = analysis()
    KPIs(R, S, sector, ticker)
    corr_ana(R, sector, rdata)
    return x, R, S

# Stock-by-Stock Trading Signals #
# These do not modify the existing API above.

# Globals used by this sub-module
SMA: pd.DataFrame | None = None
SMA_pos: pd.DataFrame | None = None
ticker: list[str] | None = None
window: list[int] | None = None

ExcelLike = Union[str, Path, IO[bytes]]

def master_data_stocktrading(
    prices_file: Optional[ExcelLike] = None,
    meta_file:   Optional[ExcelLike] = None,
    prices_sheet: str = "SP",
    meta_sheet:   Optional[str] = None,
):
    """
    Loads price & metadata for the stock-trading workflow.

    If prices_file/meta_file are omitted, this function behaves exactly like the
    original: it prompts for country and reads:
        <COUNTRY>_Share Price_Combined.xlsx  (sheet 'SP')
        1. Tic_Global.xlsx                    (sheet '<COUNTRY>')
    If file paths are provided, prompts are skipped when possible.

    Returns
    -------
    rdata, df0
    """
    try:
        # Figure out the country/sheet when using defaults
        if prices_file is None or meta_file is None or meta_sheet is None:
            _country = input("Please enter the countryname ").strip().upper()
            if not _country:
                raise ValueError("Country name cannot be empty.")
        else:
            _country = meta_sheet or ""  # may be blank if caller passed meta_sheet explicitly

        # ---- Prices workbook
        if prices_file is None:
            prices_path = f"{_country}_Share Price_Combined.xlsx"
        else:
            prices_path = prices_file

        # If it's a path-like string, check existence for friendliness
        if isinstance(prices_path, (str, Path)) and not str(prices_path).startswith("<_io"):
            if not Path(prices_path).exists():
                raise FileNotFoundError(
                    f"Could not find prices workbook: '{prices_path}'. "
                    "Provide a valid path or place it in the current folder."
                )

        xlsx_prices = pd.ExcelFile(prices_path)
        if prices_sheet not in xlsx_prices.sheet_names:
            raise KeyError(f"Sheet '{prices_sheet}' not found in prices workbook.")
        rdata = pd.read_excel(xlsx_prices, prices_sheet)
        if "Date" not in rdata.columns:
            raise KeyError("Column 'Date' missing in prices sheet.")
        rdata = rdata.set_index("Date")
        rdata = rdata.reindex(index=rdata.index[::-1])

        # ---- Meta workbook
        if meta_file is None:
            meta_path = "1. Tic_Global.xlsx"
        else:
            meta_path = meta_file

        if isinstance(meta_path, (str, Path)) and not str(meta_path).startswith("<_io"):
            if not Path(meta_path).exists():
                raise FileNotFoundError(
                    f"Could not find metadata workbook: '{meta_path}'. "
                    "Provide a valid path or place it in the current folder."
                )

        xlsx_meta = pd.ExcelFile(meta_path)

        # If caller did not specify meta_sheet explicitly, default to country sheet
        _meta_sheet = meta_sheet or _country
        if _meta_sheet not in xlsx_meta.sheet_names:
            raise KeyError(f"Sheet '{_meta_sheet}' not found in metadata workbook.")

        df0 = pd.read_excel(xlsx_meta, _meta_sheet)
        if "Company2" in df0.columns:
            df0 = df0.drop(columns=["Company2"])

        return rdata, df0

    except Exception as e:
        raise type(e)(f"[master_data_stocktrading] {e}") from e


def trading_data_stocktrading(rdata: pd.DataFrame, df0: pd.DataFrame):
    """
    Interactive part: asks for sector and three SMA windows (comma-separated).
    Returns (SMA, SMA_pos, ticker, window). Matches your notebook logic.
    """
    try:
        # sector prompt
        _sector = input("Please specify the sector you wish to focus today ? ").strip()
        if not _sector:
            raise ValueError("Sector cannot be empty.")
        print(f"Tickers of Sector: {_sector}")

        # pick tickers for sector
        if "Sector" not in df0.columns or "Ticker" not in df0.columns:
            raise KeyError("df0 must contain 'Sector' and 'Ticker' columns.")
        _tickers = list(df0.loc[df0["Sector"] == _sector, "Ticker"])
        if not _tickers:
            raise ValueError(f"No tickers found for sector '{_sector}'.")

        # windows prompt
        raw = input('Enter the three windows use "," to seperate ')
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if len(parts) < 3:
            raise ValueError("Please enter three integers, e.g. 20,42,252")
        _window = [int(p) for p in parts[:3]]
        if any(w <= 0 for w in _window):
            raise ValueError("Windows must be positive integers.")

        # compute SMA frames
        _SMA = pd.DataFrame()
        for t in _tickers:
            if t not in rdata.columns:
                # quietly skip missing tickers but keep going
                continue
            for w in _window:
                _SMA[f"{t} {int(w)}"] = rdata[t].rolling(window=int(w)).mean()

        _SMA_pos = pd.DataFrame(index=_SMA.index)
        for t in _tickers:
            if t not in rdata.columns:
                continue
            _SMA_pos[t] = rdata[t]
            _SMA_pos[f"{t} Returns"] = _SMA_pos[t].pct_change(fill_method=None)
            # positions from first two window pairs (as in your code)
            for i in range(len(_window) - 1):
                a, b = _window[i], _window[i + 1]
                a_col, b_col = f"{t} {a}", f"{t} {b}"
                if a_col in _SMA and b_col in _SMA:
                    _SMA_pos[f"Pos {t} {a}"] = np.where(_SMA[a_col] > _SMA[b_col], 1, -1)

        # stash in globals so indi_charts_stocktrading() can use them (like notebook)
        globals()["SMA"] = _SMA
        globals()["SMA_pos"] = _SMA_pos
        globals()["ticker"] = _tickers
        globals()["window"] = [int(w) for w in _window]

        return _SMA, _SMA_pos, _tickers, _window

    except Exception as e:
        raise type(e)(f"[trading_data_stocktrading] {e}") from e


def stock_data_stocktrading():
    """
    Convenience wrapper to chain the two steps, matching your notebook cell:
        rdata, df0 = master_data()
        SMA, SMA_pos, ticker, window = trading_data(rdata, df0)
    """
    rdata, df0 = master_data_stocktrading()
    _SMA, _SMA_pos, _ticker, _window = trading_data_stocktrading(rdata, df0)
    return rdata, _SMA, _SMA_pos, _ticker, _window


def indi_charts_stocktrading():
    """
    Reproduces the 2x2 summary plots loop for each ticker.
    Uses the globals created by trading_data_stocktrading()/stock_data_stocktrading().
    """
    try:
        if any(globals().get(n) is None for n in ("SMA", "SMA_pos", "ticker", "window")):
            raise RuntimeError(
                "No data to plot. Run stock_data_stocktrading() (or trading_data_stocktrading)"
                " first to populate SMA, SMA_pos, ticker, window."
            )

        _SMA: pd.DataFrame = globals()["SMA"]
        _SMA_pos: pd.DataFrame = globals()["SMA_pos"]
        _ticker: list[str] = globals()["ticker"]
        _window: list[int] = globals()["window"]

        # match your plotting style
        #from matplotlib import pyplot as plt
        #plt.style.use("seaborn-v0_8-white")
        #mpl.rcParams["font.family"] = "serif"
        #mpl.rc("xtick", labelsize=12)
        #mpl.rc("ytick", labelsize=12)

        for t in _ticker:
            if t not in _SMA_pos.columns:
                # skip tickers that were not available in rdata
                continue

            fig, ax = plt.subplots(
                2, 2, figsize=(20, 8), dpi=200, sharex=False, sharey=False,
                squeeze=False, gridspec_kw={"width_ratios": [2, 2], "height_ratios": [1, 1]},
                constrained_layout=True,
            )
            ax = ax.flatten()
            plt.rcParams.update({"font.size": 15})
            fig.suptitle(f" Summary Plots for {t}", fontsize=30, color="k", x=0.5, y=1.05)

            i = 0
            ax[i].plot(_SMA_pos[t].iloc[-100:], color="k", linewidth=3, linestyle="-", label="Stock Price")
            ax[i].set_title("Stock Price", fontsize=20, y=1.01)
            ax[i].legend()

            i += 1
            ax[i].plot(((_SMA_pos[f"{t} Returns"].iloc[-100:] + 1).cumprod() - 1),
                       color="r", linewidth=3, linestyle="--", label="Cumulative Return")
            ax[i].set_title("Cumulative Returns", fontsize=20, y=1.01)
            ax[i].legend()

            i += 1
            ma_cols = [f"{t} {_window[0]}", f"{t} {_window[1]}", f"{t} {_window[2]}"]
            present = [c for c in ma_cols if c in _SMA.columns]
            if present:
                ax[i].plot(_SMA[present].iloc[-100:])
                ax[i].legend([f"Mean {_window[0]}", f"Mean {_window[1]}", f"Mean {_window[2]}"])
            ax[i].set_title("Moving Averages", fontsize=20, y=1.01)

            i += 1
            ax[i].plot(_SMA_pos[t].iloc[-100:], color="k", linewidth=3, linestyle="-", label="Stock Price")
            ax[i].set_title("Trading Decison based on Moving Average", fontsize=20, y=1.01)
            ax[i] = ax[i].twinx()
            # same two position lines you had
            pos_a = f"Pos {t} {_window[0]}"
            pos_b = f"Pos {t} {_window[1]}"
            if pos_a in _SMA_pos.columns:
                ax[i].plot(_SMA_pos[pos_a].iloc[-100:], color="r", linewidth=3, linestyle="-",
                           label=f"{_window[0]} Vs. {_window[1]} Days")
            if pos_b in _SMA_pos.columns:
                ax[i].plot(_SMA_pos[pos_b].iloc[-100:], color="b", linewidth=3, linestyle=":",
                           label=f"{_window[1]} Vs. {_window[2]} Days")
            ax[i].legend(loc="center", bbox_to_anchor=(0.8, .3))

        plt.show()
        return

    except Exception as e:
        raise type(e)(f"[indi_charts_stocktrading] {e}") from e

# One-call interactive entry point (prompts country, sector, windows) 
def stocktrading(
    prices_file: Optional[ExcelLike] = None,
    meta_file:   Optional[ExcelLike] = None,
    prices_sheet: str = "SP",
    meta_sheet:   Optional[str] = None,
):
    """
    One-call interactive flow:
      1) If no files are given, prompts for country and uses the conventional filenames.
      2) Prompts for sector.
      3) Prompts for three SMA windows (e.g. 20,42,252).
      4) Plots the 2x2 summary charts.

    If prices_file/meta_file are provided, they will be used instead of the defaults.
    Returns (rdata, SMA, SMA_pos, ticker, window).
    """
    # Load price/meta (either via passed-in paths or the usual prompts/defaults)
    rdata, df0 = master_data_stocktrading(
        prices_file=prices_file,
        meta_file=meta_file,
        prices_sheet=prices_sheet,
        meta_sheet=meta_sheet,
    )

    # Same interactive sector + window prompts you already had
    _SMA, _SMA_pos, _ticker, _window = trading_data_stocktrading(rdata, df0)

    # Charts (unchanged)
    indi_charts_stocktrading()

    return rdata, _SMA, _SMA_pos, _ticker, _window

# end Stock-by-Stock Trading Signal #

