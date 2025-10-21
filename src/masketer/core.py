# Auto-extracted from '1. Analysis - Summary Analysis.ipynb'
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def hello():

    x = int(input(f"Which analysis you\'re looking to do today:\n\
    1. Country.\n\
    2. Sector.\n\
    ** Enter the Number from 1 or 2 **"))
    
    countryname = input(' Please name the country which you are focusing today').upper() 
            
    return x, countryname

def master_data():
    xlsx = pd.ExcelFile(f"{countryname}_Share Price_Combined.xlsx")
    rdata = pd.read_excel(xlsx, 'SP')
    rdata= rdata.set_index('Date')
    rdata = rdata.sort_values('Date').set_index('Date')
    rdata = rdata.reindex(index= rdata.index[::-1])
    # ticker = (rdata.columns.values)

    xlsx = pd.ExcelFile('1. Tic_Global.xlsx')
    df0 = pd.read_excel(xlsx, f"{countryname}") 
    #df0 = df0.drop(columns=['Company2'])
    if 'SP' not in pd.ExcelFile(...).sheet_names: raise KeyError("SP sheet missing")
    if 'Company2' in df0.columns: df0 = df0.drop(columns=['Company2'])

    pd.DataFrame().to_excel(f"{countryname}_1-Return analysis.xlsx")
    pd.DataFrame().to_excel(f"{countryname}_2-Summary analysis.xlsx")

    return rdata,df0
    

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