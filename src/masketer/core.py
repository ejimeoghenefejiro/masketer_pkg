# -------- Required for this module --------
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl  # Excel engine for .xlsx

plt.style.use("default")

import matplotlib as mpl
mpl.rcParams["font.family"] = "serif"
mpl.rc("xtick", labelsize=12)
mpl.rc("ytick", labelsize=12)

# -------- Optional / nice-to-have (import if available) --------
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


def hello():

    x = int(input(f"Which analysis you\'re looking to do today:\n\
    1. Country.\n\
    2. Sector.\n\
    ** Enter the Number from 1 or 2 **"))
    
    countryname = input(' Please name the country which you are focusing today').upper() 
            
    return x, countryname

def master_data(
    prices_file: ExcelLike | None = None,
    meta_file:   ExcelLike | None = None,
    prices_sheet: str = "SP",
    meta_sheet:   str | None = None,  # default to countryname
):
    """
    Load data either from *uploaded files* or from the default filenames.
    - prices_file: path or BytesIO to "<countryname>_Share Price_Combined.xlsx"
    - meta_file:   path or BytesIO to "1. Tic_Global.xlsx"
    - prices_sheet: sheet name for prices (default 'SP')
    - meta_sheet:   sheet name in meta workbook (default = countryname)
    Returns: rdata (Date-indexed prices), df0 (metadata)
    """
    #resolve country and sheet names
    if meta_sheet is None:
        meta_sheet = countryname  # keep your original pattern

    # ------- resolve files (uploaded or default names) -------
    # If caller supplies uploaded files (BytesIO from Streamlit/Jupyter), use them.
    # Otherwise fall back to your conventional filenames.
    if prices_file is None:
        prices_file = f"{countryname}_Share Price_Combined.xlsx"
    if meta_file is None:
        meta_file = "1. Tic_Global.xlsx"
    #safety checks (replace the invalid '...' line) -
    if not _excel_has_sheet(prices_file, prices_sheet):
        raise KeyError(f"'{prices_sheet}' sheet missing in prices workbook.")
    if not _excel_has_sheet(meta_file, meta_sheet):
        raise KeyError(f"'{meta_sheet}' sheet missing in meta workbook.")

    # read data 
    xlsx_prices = pd.ExcelFile(prices_file)
    rdata = pd.read_excel(xlsx_prices, prices_sheet)
    rdata = rdata.set_index("Date")
    rdata = rdata.sort_values("Date").set_index("Date")
    rdata = rdata.reindex(index=rdata.index[::-1])

    xlsx_meta = pd.ExcelFile(meta_file)
    df0 = pd.read_excel(xlsx_meta, meta_sheet)

    # optional cleanup, only if present (keeps your original intent)
    if "Company2" in df0.columns:
        df0 = df0.drop(columns=["Company2"])

    # ------- create placeholders as before -------
    pd.DataFrame().to_excel(f"{countryname}_1-Return analysis.xlsx")
    pd.DataFrame().to_excel(f"{countryname}_2-Summary analysis.xlsx")

    return rdata, df0

    

def country_analysis(countryname):

    R= rdata.pct_change(fill_method=None)
    R.to_excel(f"{countryname}_return analysis.xlsx")

    S=pd.DataFrame(R.describe(percentiles=[0.01,0.1,0.5,0.9,0.99]))
    S.to_excel(f"{countryname}_summary analysis.xlsx")
    
    return R,S

def sector_analysis(countryname):   #exisitng dataset for the entire country is available
    sector = input('Please specify the sector you wish to focus today ? ')
    print(f"Tickers of Sector: {sector}")  
    ticker= list(df0.loc[df0['Sector'] == sector,'Ticker'])

#--------------------------------------------------------------------------#

    #---STEP 1 RETURN---
    
    R= rdata[[x for x in ticker]].pct_change(fill_method = None)
  
    
    #---STEP 2 COMBINE NEW AND OLD RETURN---

    # with pd.ExcelWriter(f"{countryname}_1-Return analysis.xlsx", engine='openpyxl', mode='a',if_sheet_exists = 'replace') as writer:  
    #     R.to_excel(writer, sheet_name=f"{sector[:20]}_{countryname}")

    #---STEP 3 COMBINE NEW AND OLD SUMMARY---

    S=pd.DataFrame(R.describe(percentiles=[0.01,0.1,0.5,0.9,0.99]))

    # with pd.ExcelWriter(f"{countryname}_2-Summary analysis.xlsx", engine='openpyxl', mode='a',if_sheet_exists = 'replace') as writer:  
    #     S.to_excel(writer, sheet_name=f"{sector[:20]}_{countryname}")
    
    return sector,ticker,R,S

###----------------------------------------------------------------------------------------------------------------------###
def KPIs(R,S,sector,ticker):  

    # R,S,rdata =  dataset(rdata)
    
## bar chart cumulative return
    
    
    a= (((R+1).cumprod())-1).iloc[-2:-1].round(3)
    
    a.T.plot(kind='bar',color='k',figsize=(12,4),title='Distribution of Cumulative Returns',fontsize=10, grid=True,legend=False)

    plt.savefig(f"cum_ret_dis_{sector}_{countryname}", bbox_inches='tight')

    

## Graphs of Summary
    
    
    plt.figure(figsize=(12,15))
    plt.suptitle("Stocks' KPIS for Investors",fontsize=16, color='r',y=0.93)
    
    plt.subplot(221)
    plt.title('Average Return on Stocks (Daily)',fontsize=14, y=1.01)
    S.loc['mean'].sort_values(ascending=False).plot(kind='barh',fontsize=10,
                                                    color=['r','g','b','c','k'],edgecolor='k',linestyle='--',grid=True)
    
    plt.subplot(222)
    plt.title('Average Voltality in Stock Returns (Daily)',fontsize=14 ,y=1.01)
    S.loc['std'].sort_values(ascending=False).plot(kind='barh',fontsize=10,
                                                   color=['r','g','b','c','k'],edgecolor='k',linestyle='--',grid=True)
    
    plt.savefig(f"risk_ret_{sector}_{countryname}", bbox_inches='tight')
    plt.show()
# average cumulatiev daily return

    S.loc[['1%','10%','50%','90%','99%']].plot(kind='box',vert=True, figsize=(12,4),title='Distribution of Daily Returns',fontsize=10,grid=True,
                                              color='k',style='-b',widths=0.1,showcaps=True,showbox=True,showmeans=True,
                                               boxprops=dict(linestyle='-', linewidth=1.5))
    plt.savefig(f"high_low_ret_{sector}_{countryname}", bbox_inches='tight')
    return