# -*- coding: gbk -*-
from standard_product_data import *
import pandas as pd
import utilities as ut
from copy import deepcopy

#accounts账目的类
class Asset:#基类
    def __init__(self, StandProduct):
        #所有调用都用列名不要用列序号，即用.loc
        self.StandProduct = StandProduct
        self.classify_holding()
        self.money_fund = self.calculate_money_fund()
        self.money_fund_trade_value = self.calculate_money_fund_trade_value()
        self.stock_holding_value_today = self.calculate_stock_holding_value_today()
        self.stock_value_yesterday = self.calculate_stock_value_yesterday()
        self.stock_value_today = self.calculate_stock_value_today()
        if self.StandProduct.product_id in [0, 1, 2, 4, 5, 7, 11, 12, 13, 16, 17, 19, 101, 102, 107, 208]:
            self.total_settle, self.total_closing = self.fut_value_today()
        else:
            self.total_settle, self.total_closing = 0, 0
        # print self.total_settle, self.total_closing
        self.calculate_static_rights_and_interest()
        self.trade_cost = self.calculate_trade_cost()
        self.trade_fut_cost = self.calculate_trade_fut_cost()
        self.trade_commission, self.trade_fut_commission = self.calculate_commission()
        # print self.trade_cost
        # print self.trade_fut_cost
        # print self.holding_fut_pre_d
        # print self.trade_fut_commission
        # print self.trade_fut_cost+self.holding_fut_pre_d+self.trade_fut_commission
        self.buyback = self.calculate_buyback()
        self.total_buy, self.total_sell = self.calculate_trade_total()
        self.total_long_value, self.total_short_value = self.calculate_future_total()
        self.compare_return_and_holding_price()

    def classify_holding(self):
        holding_df = self.StandProduct.csv['holding']
        today_date = self.StandProduct.t.date
        df = ut.classify_holding_type(holding_df, today_date)

        df1 = df[df['H'] == 1]
        self.holding_Hushi = sum(df1['Holding'] * df1['Price'])

        df1 = df[df['S'] == 1]
        self.holding_Shenshi = sum(df1['Holding'] * df1['Price'])

        df1 = df[df['sh50'] == 1]
        self.holding_sh50 = sum(df1['Holding'] * df1['Price'])

        df1 = df[df['hs300'] == 1]
        self.holding_hs300 = sum(df1['Holding'] * df1['Price'])

        df1 = df[df['csi500'] == 1]
        self.holding_csi500 = sum(df1['Holding'] * df1['Price'])




    def calculate_money_fund(self):
        holding_df = self.StandProduct.csv['holding']
        holding_df = holding_df[holding_df['Commission_type'] != 0.7]
        holding_df = holding_df[holding_df['Commission_type'] != 1]
        return sum(holding_df['Holding'] * holding_df['Price'])

    def calculate_money_fund_trade_value(self):
        trade_df = self.StandProduct.csv['trade']
        trade_df = trade_df[trade_df['Commission_type'] != 0.7]
        trade_df = trade_df[trade_df['Commission_type'] != 1]
        trade_df['Dir_volume'] = trade_df['Direction'].apply(lambda x: 1 if x == 'buy' else -1) * trade_df['Volume']
        return sum(trade_df['Dir_volume'] * trade_df['Price'])

    def calculate_stock_holding_value_today(self):      # 用holding  里面的价格
        holding_df = self.StandProduct.csv['holding']
        holding_df = holding_df[holding_df['Commission_type'] == 1]
        stock_holding_value_today = sum(holding_df['Holding'] * holding_df['Price'])
        print 'stock_holding_value_today', stock_holding_value_today
        return sum(holding_df['Holding'] * holding_df['Price'])

    def calculate_stock_value_yesterday(self):
        holding_df = self.StandProduct.csv['holding']
        holding_df = holding_df[holding_df['Commission_type'] == 1]
        trade_df = self.StandProduct.csv['trade']
        trade_df = trade_df[trade_df['Commission_type'] == 1]
        print holding_df[holding_df['str_id'] == '600020']

        print trade_df[trade_df['Id'] == 600020]

        holding_pre_df = deepcopy(holding_df)
        if not trade_df.empty:
            trade_df = trade_df.reset_index(drop = True)
            trade_df['Dir_volume'] = trade_df['Direction'].apply(lambda x: 1 if x=='buy' else -1) * trade_df['Volume']
            temp = trade_df.groupby('Id').sum().loc[:,['Dir_volume']].reset_index()
            m = pd.merge(holding_pre_df, temp, how='outer', on='Id').fillna(0)
            m['Holding'] = m['Holding'] - m['Dir_volume']
        else:
            m = holding_pre_df

        print m

        return_td = self.StandProduct.t.stock_return
        m = pd.merge(m, return_td, left_on='Id', right_on='id').sort_values(by='Id')
        print 'stock_value_yesterday', sum(m['Holding'] * m['px_last'])
        return sum(m['Holding'] * m['px_last'])

    def calculate_stock_value_today(self):      # 用return里面的价格
        holding_df = self.StandProduct.csv['holding']
        holding_df = holding_df[holding_df['Commission_type'] == 1]
        return_td = self.StandProduct.t.stock_return
        m = pd.merge(holding_df, return_td, left_on='Id', right_on='id').sort_values(by='Id')
        print 'stock_value_today', sum(m['Holding'] * m['px_last'])
        return sum(m['Holding'] * m['px_last'])

    def fut_value_yesterday(self):#需要今天的市值, 昨天的holding
        total = 0.0
        return total

    def fut_value_today(self):#需要今天的市值, 今天的holding
        total_settle = 0.0
        total_closing = 0.0
        today_date = self.StandProduct.t.date
        futpx = self.StandProduct.t.futpx
        futpx_today = futpx[futpx['date'] == today_date].copy()
        holding_fut_df = self.StandProduct.csv['holding_fut']
        for i in range(len(holding_fut_df)):
            holding_fut_row = holding_fut_df.ix[i]
            fut_type = holding_fut_row['Id'][:2]
            holding = holding_fut_row['Holding']
            direction = holding_fut_row['Direction']
            m = futpx_today[futpx_today['type'] == fut_type].copy()
            m = m[m['contract_name'] == holding_fut_row['Contract_name']].copy()

            price_settle = float(m['settle'])
            price_closing = float(m['close'])
            multiply = int(m['multiply'])
            if direction == 'buy':
                total_settle = total_settle + price_settle * holding * multiply
                total_closing = total_closing + price_closing * holding * multiply
            else:
                total_settle = total_settle - price_settle * holding * multiply
                total_closing = total_closing - price_closing * holding * multiply
        print 'total_settle',total_settle   #结算价
        print 'total_closing', total_closing  #收盘价
        return total_settle,total_closing

    def calculate_static_rights_and_interest(self):
        futpx = self.StandProduct.t.futpx
        today_date = self.StandProduct.t.date
        yesterday = self.StandProduct.t.yesterday
        futpx_today = futpx[futpx['date'] == today_date].copy()
        futpx_yesterday = futpx[futpx['date'] == yesterday].copy()
        futpx_today['ID'] = futpx_today['id'].apply(lambda x: x[:2])
        futpx_yesterday['ID'] = futpx_yesterday['id'].apply(lambda x: x[:2])
        holding_fut_pre_df = self.StandProduct.csv['holding_fut_pre']
        holding_fut_pre_df['ID'] = holding_fut_pre_df['Id'].apply(lambda x: x[:2])
        m = pd.merge(holding_fut_pre_df, futpx_yesterday, left_on=['ID', 'Contract_name'], right_on=['ID', 'contract_name'])
        m = pd.merge(m, futpx_today, left_on=['ID', 'Contract_name'], right_on=['ID', 'contract_name'])
        m['type'] = m['Direction'].apply(lambda x: 1 if x == 'buy' else -1)
        self.holding_fut_pre_d = sum((m['settle_y'] - m['settle_x']) * m['Holding'] * m['type'] * m['multiply_x'])

        holding_pre_df = self.StandProduct.csv['holding_pre']
        holding_pre_df = holding_pre_df[holding_pre_df['Commission_type'] == 1]
        return_pre = self.StandProduct.t.stock_pre_return.loc[:, ['id', 'px_last']]

        return_td = self.StandProduct.t.stock_return.loc[:, ['id', 'ret']]

        m = pd.merge(holding_pre_df, return_pre, left_on='Id', right_on='id', how='left')
        m = pd.merge(m, return_td, left_on='Id', right_on='id', how='left')
        self.holding_pre_d = sum(m['Holding'] * m['px_last'] * m['ret'])

    def calculate_trade_cost(self):
        trading_price = self.StandProduct.csv['trade']
        trading_price = trading_price[trading_price['Commission_type'] == 1]

        if not trading_price.empty:
            real_price = self.StandProduct.t.stock_return
            m = pd.merge(trading_price, real_price, how='left', left_on='Id', right_on='id')
            m['cost'] = m.apply(lambda x: (1 if x['Direction']=='buy' else -1)*(x['px_last']-x['Price'])*x['Volume'], axis=1)
            print 'trade_cost', sum(m['cost'])
            return sum(m['cost'])
        else:
            return 0.0

    def calculate_trade_fut_cost(self):
        cost = 0.0
        trading_price = self.StandProduct.csv['trade_fut']
        today_date = self.StandProduct.t.date
        futpx = self.StandProduct.t.futpx
        futpx_today = futpx[futpx['date'] == today_date].copy()
        for i in range(len(trading_price)):
            trading_price_row = trading_price.ix[i]
            fut_type = trading_price_row['Id'][:2]
            volume = trading_price_row['Volume']
            price = float(trading_price_row['Price'])

            m = futpx_today[futpx_today['type'] == fut_type].copy()
            m = m[m['contract_name'] == trading_price_row['Contract_name']].copy()

            px_last = float(m['settle'])#.astype(np.float)
            multiply = int(m['multiply'])
            if trading_price_row['Direction'] == 'buy':
                cost = cost + (px_last - price) * volume * multiply
            else:
                cost = cost - (px_last - price) * volume * multiply
        print 'trade_fut_cost',cost
        return cost

    def calculate_commission(self):
        trade_df = self.StandProduct.csv['trade']
        if not trade_df.empty:
            trade_df['Commission_after'] = trade_df.apply(lambda x: (1 if x['Commission_type'] == 0.5 else x['Commission_type']) * x['Commission_cal'], axis=1)
            trade_commission = sum(trade_df['Commission_after'])
        else:
            trade_commission = 0
        trade_fut_df = self.StandProduct.csv['trade_fut']
        trade_fut_commission = sum(trade_fut_df['Commission'])
        print 'trade_commission',-trade_commission
        print 'trade_fut_commission',-trade_fut_commission
        return -trade_commission, -trade_fut_commission


    def calculate_buyback(self):
        holding_df = self.StandProduct.csv['holding']
        holding_df = holding_df[holding_df['Commission_type']==0.7]
        print 'Buyback',sum(holding_df['Holding'] * holding_df['Price'])
        return sum(holding_df['Holding'] * holding_df['Price'])


    def calculate_trade_total(self):
        trade_df = self.StandProduct.csv['trade']
        if (len(trade_df[trade_df['Direction']=='buy'])!=0):
            total_buy = sum(trade_df[trade_df['Direction']=='buy'].apply(lambda x: x['Price']*x['Volume'], axis=1))
        else:
            total_buy = 0
        if (len(trade_df[trade_df['Direction']=='sell'])!=0):
            total_sell = sum(trade_df[trade_df['Direction']=='sell'].apply(lambda x: x['Price']*x['Volume'], axis=1))
        else:
            total_sell = 0
        print 'total buy',total_buy
        print 'total sell',total_sell
        return total_buy, total_sell

    def calculate_future_total(self):
        total_long_value = 0.0
        total_short_value = 0.0
        today_date = self.StandProduct.t.date
        futpx = self.StandProduct.t.futpx
        futpx_today = futpx[futpx['date'] == today_date].copy()
        holding_fut_df = self.StandProduct.csv['holding_fut']
        for i in range(len(holding_fut_df)):
            holding_fut_row = holding_fut_df.ix[i]
            fut_type = holding_fut_row['Id'][:2]
            holding = holding_fut_row['Holding']
            m = futpx_today[futpx_today['type'] == fut_type].copy()
            m = m[m['contract_name'] == holding_fut_row['Contract_name']].copy()
            settle_price = float(m['settle'])
            multiply = int(m['multiply'])
            if holding_fut_row['Direction'] == 'buy':
                total_long_value = total_long_value + settle_price * holding * multiply
            else:
                total_short_value = total_short_value + settle_price * holding * multiply
        print 'total_long_value', total_long_value
        print 'total_short_value', total_short_value
        return total_long_value, total_short_value

    def compare_return_and_holding_price(self):
        holding_df = self.StandProduct.csv['holding'].loc[:, ['Id', 'Price']]
        return_td = self.StandProduct.t.stock_return.loc[:, ['id', 'px_last']]
        m = pd.merge(holding_df, return_td, left_on = 'Id', right_on = 'id')
        df = m[m['px_last'] != m['Price']]
        if df.empty:
            print ur'return.csv 与 holding 股票价格一致'
        else:
            print ur'return.csv 与 holding 股票价格对比结果如下'
            print df

    def trade_today(self):
        return

    def trading_fee(self):
        return

    def trading_gain_loss(self):
        return

    def profit(self):
        self.profit = self.value_today() - self.trade_today() - self.value_yesterday()
        return self.profit

    def ret(self):
        self.ret = self.profit()/self.value_yesterday()
        return self.ret

class Stock(Asset):
    def __init__(self, standard_product_data):
        self.standard = standard_product_data


class FixedIncome(Asset): #货币基金、回购等固定收益
    holding_today = []#从holding.csv里面筛选出来
    value_today = 0.0
    holding_yesterday = []
    value_yesterday = 0.0


class Future(Asset):
    pass

class Commodity:
    pass

class Liabilities:
    def __init__(self):
        self.liabilities = {}
        self.liabilities['glrbc'] = self.glrbc()
        self.liabilities['zctgf'] = self.zctgf()
        self.liabilities['tzgwf'] = self.tzgwf()

        self.total_liabilities = sum(self.liabilities.values())
    pass

class Accounts:
    def __init__(self):
        self.s = Stock()
        self.f = Future()
        self.c = Commodity()
        self.config = pd.read_csv("{}/python/common_data/accounts_config.csv".format(t.root_path))

    def beta(self): #既有股票、又有期货的风控数据写在这个类下面
        pass

    def near_month_net_value(self):
        pass

    def seasonal_net_annualized_ret(self):
        pass

