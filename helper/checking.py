import pandas as pd
from pandas.tseries.offsets import MonthEnd
import os
import time

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
pd.set_option('float_format', '{:f}'.format)

def Checking(filePath, startdate):
    global aum
    aum = pd.read_excel(filePath, sheet_name='Data', header=1)
    aum = aum[['Monthend', 'Fund', 'Fund_Series', 'NAV', 'NAV_ps', 'Commitment', 'Subscription / Drawdown', 'Redemption', 'Distribution', 'Committed Capital', 'Total No. of shr','Monthly MF', 'o/s accrued PF']]
    aum = aum.rename(columns={'Monthend': 'date', 'Fund': 'fund', 'Fund_Series': 'series', 'NAV': 'nav', 'NAV_ps': 'navps', 'Commitment': 'commitment', 'Subscription / Drawdown': 'subscription', 'Redemption' :'redemption', 'Distribution': 'distribution', 'Committed Capital': 'committed', 'Total No. of shr': 'shares', 'Monthly MF': 'mf', 'o/s accrued PF': 'pf'})
    aum['series'] = aum['series'].fillna('series0')
    aum = aum.query('date > @startdate')    # check what date



    aum['date'] = pd.to_datetime(aum['date'], format="%Y%m") + MonthEnd(0)
    aum[['date']] = aum[['date']].astype(str)

    aum = aum.groupby(['date', 'fund', 'series']).sum()
    aum = aum.sort_values(by=['fund', 'series', 'date'])
    # aum = aum.reset_index(level=0)
    #aum[['date']] = aum[['date']].dt.date
    #aum.dates.apply(lambda x: x.date())
    #aum['date'] = [date.date() for date in aum['date']]

    aum[['nav_chg', 'navps_chg', 'commit_chg', 'sub_chg', 'redem_chg', 'dis_chg', 'commited_chg', 'shrs_chg', 'mf_chg',
         'pf_chg']] = aum.groupby(['fund', 'series'])[
        ['nav', 'navps', 'commitment', 'subscription', 'redemption', 'distribution', 'committed', 'shares', 'mf',
         'pf']].apply(lambda x: x.pct_change())

    aum = aum[
        ['nav', 'nav_chg', 'navps', 'navps_chg', 'commitment', 'commit_chg', 'subscription', 'sub_chg', 'redemption',
         'redem_chg', 'distribution', 'dis_chg', 'committed', 'commited_chg', 'shares', 'shrs_chg', 'mf', 'mf_chg',
         'pf', 'pf_chg']]

    aum = aum.style.applymap(check_3_pct,subset=['nav_chg', 'navps_chg', 'shrs_chg', 'mf_chg', 'pf_chg']).format({'nav':"{:.2f}",'navps':"{:.2f}"
        ,'commitment': "{:.2f}", 'subscription': "{:.2f}", 'redemption': "{:.2f}", 'distribution': "{:.2f}", 'committed': "{:.2f}"})

    return aum

def download_excel():
    # writer = pd.ExcelWriter('C:/Users/kellyshum/Downloads/ABCD.xlsx', engine='openpyxl')

    timestr = time.strftime("%Y%m%d-%H%M%S")
    aum.to_excel('result %s.xlsx' % (timestr))
    os.startfile('result %s.xlsx' % (timestr))
    return aum


def check_3_pct(val):
    color_val = ''
    if val > 0.03 or val < -0.03:
        color_val = 'yellow'
    return 'background-color: %s' % color_val
