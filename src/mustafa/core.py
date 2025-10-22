#Required for this module 
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl  # Excel engine for .xlsx
from pathlib import Path
from typing import Union, IO, Optional

plt.style.use("default")

import matplotlib as mpl
mpl.rcParams["font.family"] = "serif"
mpl.rc("xtick", labelsize=12)
mpl.rc("ytick", labelsize=12)



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
