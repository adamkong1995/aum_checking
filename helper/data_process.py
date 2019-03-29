from helper.config import pd
from pandas.tseries.offsets import MonthEnd
import datetime
import dateutil.relativedelta
from helper.checking import to_check


def process(filePath, startdate):
    global aum
    aum = pd.read_excel(filePath, sheet_name='Data', header=1)
    aum = aum[['Monthend', 'Fund', 'Fund_Series', 'NAV', 'NAV_ps', 'Commitment', 'Subscription / Drawdown', 'Redemption', 'Distribution', 'Committed Capital', 'Total No. of shr','Monthly MF', 'o/s accrued PF']]
    aum = aum.rename(columns={'Monthend': 'date', 'Fund': 'fund', 'Fund_Series': 'series', 'NAV': 'nav', 'NAV_ps': 'navps', 'Commitment': 'commitment', 'Subscription / Drawdown': 'subscription', 'Redemption' :'redemption', 'Distribution': 'distribution', 'Committed Capital': 'committed', 'Total No. of shr': 'shares', 'Monthly MF': 'mf', 'o/s accrued PF': 'pf'})
    aum['series'] = aum['series'].fillna('series0')

    # month minus 1
    d = datetime.datetime.strptime(startdate, "%m-%Y")
    d2 = d - dateutil.relativedelta.relativedelta(months=1)

    aum = aum.query('date > @d2')    # check what date

    aum['date'] = pd.to_datetime(aum['date'], format="%Y%m") + MonthEnd(n=0)
    aum[['date']] = aum[['date']].astype(str)

    aum = aum.groupby(['date', 'fund', 'series']).sum()
    aum = aum.sort_values(by=['fund', 'series', 'date'])

    aum[['nav_chg', 'navps_chg', 'commit_chg', 'sub_chg', 'redem_chg', 'dis_chg', 'commited_chg', 'shrs_chg', 'mf_chg',
         'pf_chg']] = aum.groupby(['fund', 'series'])[
        ['nav', 'navps', 'commitment', 'subscription', 'redemption', 'distribution', 'committed', 'shares', 'mf',
         'pf']].apply(lambda x: x.pct_change())

    aum = aum[
        ['nav', 'nav_chg', 'navps', 'navps_chg', 'commitment', 'commit_chg', 'subscription', 'sub_chg', 'redemption',
         'redem_chg', 'distribution', 'dis_chg', 'committed', 'commited_chg', 'shares', 'shrs_chg', 'mf', 'mf_chg',
         'pf', 'pf_chg']]

    aum = aum.style.applymap(to_check, subset=['nav_chg', 'navps_chg', 'shrs_chg', 'mf_chg', 'pf_chg']).format({'nav':"{:.2f}", 'navps':"{:.2f}", 'commitment': "{:.2f}", 'subscription': "{:.2f}", 'redemption': "{:.2f}", 'distribution': "{:.2f}", 'committed': "{:.2f}"})

    return aum

def download_excelfile():
    excel = aum.to_excel('./download_folder/result.xlsx')
    return excel
