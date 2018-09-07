# -*- coding: gbk -*-
from standard_product_data import *
import pandas as pd

#策略的类

class TotalAlpha():
    def __init__(self, StandProduct):
        self.StandProduct = StandProduct
        self.calculate_premium()
        self.calculate_total_mix_return()
        self.calculate_target_Totalret()
        self.calculate_holding_raw_ret()
        self.calculate_holding_beta()

    def calculate_total_mix_return(self):
        self.index_hedge_target = self.StandProduct.index_hedge_target


        x = -sum(self.index_hedge_target[self.index_hedge_target['long'] == 'IF']['long_amount'])
        y = -sum(self.index_hedge_target[self.index_hedge_target['long'] == 'IH']['long_amount'])
        z = -sum(self.index_hedge_target[self.index_hedge_target['long'] == 'IC']['long_amount'])
        futpx_today = self.futpx_today
        futpx_yesterday = self.futpx_yesterday


        futpx_x_today = futpx_today[futpx_today['type'] == 'IF'].copy()
        futpx_x_today = futpx_x_today[futpx_x_today['contract_name'] == 'current'].copy()
        futpx_y_today = futpx_today[futpx_today['type'] == 'IH'].copy()
        futpx_y_today = futpx_y_today[futpx_y_today['contract_name'] == 'current'].copy()
        futpx_z_today = futpx_today[futpx_today['type'] == 'IC'].copy()
        futpx_z_today = futpx_z_today[futpx_z_today['contract_name'] == 'current'].copy()

        futpx_x_yesterday = futpx_yesterday[futpx_yesterday['type'] == 'IF'].copy()
        futpx_x_yesterday = futpx_x_yesterday[futpx_x_yesterday['contract_name'] == 'current'].copy()
        futpx_y_yesterday = futpx_yesterday[futpx_yesterday['type'] == 'IH'].copy()
        futpx_y_yesterday = futpx_y_yesterday[futpx_y_yesterday['contract_name'] == 'current'].copy()
        futpx_z_yesterday = futpx_yesterday[futpx_yesterday['type'] == 'IC'].copy()
        futpx_z_yesterday = futpx_z_yesterday[futpx_z_yesterday['contract_name'] == 'current'].copy()


        x_pre_value = x * float(futpx_x_yesterday['target']) * float(futpx_x_yesterday['multiply'])
        y_pre_value = y * float(futpx_y_yesterday['target']) * float(futpx_y_yesterday['multiply'])
        z_pre_value = z * float(futpx_z_yesterday['target']) * float(futpx_z_yesterday['multiply'])
        x_td_value = x * float(futpx_x_today['target']) * float(futpx_x_today['multiply'])
        y_td_value = y * float(futpx_y_today['target']) * float(futpx_y_today['multiply'])
        z_td_value = z * float(futpx_z_today['target']) * float(futpx_z_today['multiply'])
        self.index_pre_value = x_pre_value + y_pre_value + z_pre_value
        self.index_td_value = x_td_value + y_td_value + z_td_value
        self.mix_ret = float(self.index_td_value) / self.index_pre_value - 1 if self.index_pre_value else 0
        print "index_pre_value\t\t{}\nindex_td_value\t\t{}\nmix_ret\t\t{}".format(self.index_pre_value, self.index_td_value, self.mix_ret)

    def calculate_premium(self):
        futpx = self.StandProduct.t.futpx
        today_date = self.StandProduct.t.date
        yesterday_date = self.StandProduct.t.yesterday
        self.futpx_today = futpx[futpx['date'] == today_date].copy()
        self.futpx_yesterday = futpx[futpx['date'] == yesterday_date].copy()
        self.all_future = self.StandProduct.all_future
        # print 'self.all_future',self.all_future
        # print(self.all_future)
        for a in self.all_future.values:
            if a[2]:
                self.all_future = self.all_future.append(pd.DataFrame([[a[2], -a[3], '', 0, a[4], a[6], '']],columns=['long', 'long_amount', 'short', 'short_amount', 'property','long_contract', 'short_contract']))

        self.all_future = self.all_future[self.all_future['long'] != '']
        self.all_future['contract_name'] = self.all_future['long_contract'].apply(lambda x: self.StandProduct.month2contract(x[2:]))
        self.index_hedge_premium = self.all_future[self.all_future['property'] != 1]  # 只算跨指数的部分，这里排除了单边1
        self.beta_index = self.StandProduct.beta_index
        # print(self.beta_index)
        self.cross_index = self.all_future[self.all_future['property'] == 2]  # 跨指数的部分
        self.single_side = self.all_future[self.all_future['property'] == 1]  # 单边的部分
        m = pd.merge(self.futpx_today, self.index_hedge_premium, how = 'right', right_on = ['long', 'contract_name'], left_on = ['type', 'contract_name'])
        m = pd.merge(self.futpx_yesterday, m, how = 'right', right_on = ['long', 'contract_name'], left_on = ['type','contract_name'])
        # m = m.loc[:,['premium_today','premium_yesterday', 'shsz', 'long', 'long_amount', 'multiply']]
        if not len(m):
            print "No premium"
            self.premium = 0
            return
        m['earning'] = m.apply(lambda x: x['settle_y'] - x['settle_x'] + x['target_x'] - x['target_y'],axis=1)  # x,y分别代表昨天和今天
        self.premium = sum(m['long_amount'] * m['earning'] * m['multiply_x'])
        print "premium_fluctuation\t\t{}".format(self.premium)


    def calculate_target_Totalret(self, strategy='final'):
        target_pre = self.StandProduct.stock_target_pre[strategy].loc[:, ['coid', 'stockhand']]
        self.return_td = self.StandProduct.t.stock_return.loc[:, ['id', 'ret']].fillna(value=0)
        self.return_pre = self.StandProduct.t.stock_pre_return.loc[:, ['id', 'px_last']].fillna(value=0)
        m = pd.merge(target_pre, self.return_td, how='left', left_on='coid', right_on='id').sort_values(by='coid')
        m = pd.merge(m, self.return_pre, how='left', left_on='coid', right_on='id').sort_values(by='coid')
        m = m[m['stockhand'] != 0]
        self.target_pre_value = sum(m['stockhand'] * m['px_last'] * 100)  # 股票的前一交易日收盘价格的市值
        self.target_raw_ret = sum(m['stockhand'] * m['ret'] * m['px_last'] * 100) / (0.00000000001 + self.target_pre_value)  # 股票的今
        self.beta = 1
        print 'total_target_mix_ret', self.mix_ret
        self.target_real_ret = self.target_raw_ret - self.beta * self.mix_ret
        self.StandProduct.target_real_ret_dict[strategy] = self.target_real_ret
        print self.StandProduct.target_real_ret_dict
        self.StandProduct.target_pre_value_dict[strategy] = self.target_pre_value
        print strategy,"target_pre_value\t\t{}\ntarget_raw_ret\t\t{}".format(self.target_pre_value,self.target_raw_ret)  # 总的target 昨日收盘市值,收益率



    def calculate_holding_raw_ret(self):
        df = self.StandProduct.csv["holding_pre"][self.StandProduct.csv["holding_pre"]['Commission_type'] == 1].copy()
        holding_pre = df.loc[:,['Id','Holding']]
        self.return_td = self.StandProduct.t.stock_return.loc[:, ['id', 'ret']].fillna(value=0)
        self.return_pre = self.StandProduct.t.stock_pre_return.loc[:, ['id', 'px_last']].fillna(value=0)
        m = pd.merge(holding_pre, self.return_td, left_on = 'Id', right_on = 'id').sort_values(by = 'Id')
        # m = pd.merge(m, self.return_pre, how='left', left_on='Id', right_on='id').sort_values(by='Id')
        m = pd.merge(m, self.return_pre, left_on='Id', right_on='id').sort_values(by='Id')
        self.holding_pre_value = sum(m['Holding'] * m['px_last'])  # 股票的前一交易日收盘价格的市值
        self.holding_raw_ret = sum(m['Holding'] * m['ret'] * m['px_last']) / (0.00000000001 + self.holding_pre_value)  # 股票的昨日holding，今日市值
        print "holding_pre_value\t\t{}\nholding_raw_ret\t\t{}".format(self.holding_pre_value, self.holding_raw_ret)  # 总的昨日的holding 昨日收盘市值,收益率


    def calculate_holding_beta(self):
        self.holding_beta = float(self.index_pre_value) / self.holding_pre_value if self.holding_pre_value else 0
        print 'index_pre_value', self.index_pre_value
        print 'holding_pre_value', self.holding_pre_value
        print 'holding_beta', self.holding_beta


class Alpha():
    def __init__(self, StandProduct, TotalAlpha):
        self.StandProduct = StandProduct
        self.TotalAlpha = TotalAlpha
        self.calculate_mix_return()
        self.calculate_target_real_ret()

    def calculate_mix_return(self, strategy='dvd_uncon_con50'):
        if self.StandProduct.product_id in [12]:
            target_pre = self.StandProduct.future_target_pre['dvd_con50']
        else:
            if self.StandProduct.product_id in [17]:
                target_pre = self.StandProduct.future_target_pre['port_intraday50']
            else:
                target_pre = self.StandProduct.future_target_pre[strategy]
        x = float(target_pre['future_number_300'])
        y = float(target_pre['future_number_50'])
        z = float(target_pre['future_number_500'])

        futpx_today = self.TotalAlpha.futpx_today
        futpx_yesterday = self.TotalAlpha.futpx_yesterday
        futpx_x_today = futpx_today[futpx_today['type'] == 'IF'].copy()
        futpx_x_today = futpx_x_today[futpx_x_today['contract_name'] == 'current'].copy()
        futpx_y_today = futpx_today[futpx_today['type'] == 'IH'].copy()
        futpx_y_today = futpx_y_today[futpx_y_today['contract_name'] == 'current'].copy()
        futpx_z_today = futpx_today[futpx_today['type'] == 'IC'].copy()
        futpx_z_today = futpx_z_today[futpx_z_today['contract_name'] == 'current'].copy()

        futpx_x_yesterday = futpx_yesterday[futpx_yesterday['type'] == 'IF'].copy()
        futpx_x_yesterday = futpx_x_yesterday[futpx_x_yesterday['contract_name'] == 'current'].copy()
        futpx_y_yesterday = futpx_yesterday[futpx_yesterday['type'] == 'IH'].copy()
        futpx_y_yesterday = futpx_y_yesterday[futpx_y_yesterday['contract_name'] == 'current'].copy()
        futpx_z_yesterday = futpx_yesterday[futpx_yesterday['type'] == 'IC'].copy()
        futpx_z_yesterday = futpx_z_yesterday[futpx_z_yesterday['contract_name'] == 'current'].copy()


        x_pre_value = x * float(futpx_x_yesterday['target']) * float(futpx_x_yesterday['multiply'])
        y_pre_value = y * float(futpx_y_yesterday['target']) * float(futpx_y_yesterday['multiply'])
        z_pre_value = z * float(futpx_z_yesterday['target']) * float(futpx_z_yesterday['multiply'])
        x_td_value = x * float(futpx_x_today['target']) * float(futpx_x_today['multiply'])
        y_td_value = y * float(futpx_y_today['target']) * float(futpx_y_today['multiply'])
        z_td_value = z * float(futpx_z_today['target']) * float(futpx_z_today['multiply'])
        self.index_pre_value = x_pre_value + y_pre_value + z_pre_value
        self.index_td_value = x_td_value + y_td_value + z_td_value
        self.mix_ret = float(self.index_td_value) / self.index_pre_value - 1 if self.index_pre_value else 0

    def calculate_target_real_ret(self, strategy='dvd_uncon_con50'):
        if self.StandProduct.product_id in [12]:
            target_pre = self.StandProduct.stock_target_pre['dvd_con50']
        else:
            if self.StandProduct.product_id in [17]:
                target_pre = self.StandProduct.stock_target_pre['port_intraday50']
            else:
                target_pre = self.StandProduct.stock_target_pre[strategy]

        self.return_td = self.StandProduct.t.stock_return.loc[:, ['id', 'ret']].fillna(value=0)
        self.return_pre = self.StandProduct.t.stock_pre_return.loc[:, ['id', 'px_last']].fillna(value=0)
        m = pd.merge(target_pre, self.return_td, how='left', left_on='coid', right_on='id').sort_values(by='coid')
        m = pd.merge(m, self.return_pre, how='left', left_on='coid', right_on='id').sort_values(by='coid')
        self.target_pre_value = sum(m['stockhand'] * m['px_last'] * 100)  # 股票的前一交易日收盘价格的市值
        self.target_raw_ret = sum(m['stockhand'] * m['ret'] * m['px_last'] * 100) / (0.00000000001 + self.target_pre_value)  # 股票的今日
        print strategy,"{} target_pre_value\t\t{}\n{} target_raw_ret\t\t{}".format(strategy, self.target_pre_value,strategy,self.target_raw_ret)  # 每个策略的 昨日收盘市值,收益率
        self.beta = 1
        print  strategy, 'mix_ret', self.mix_ret
        self.real_ret = self.target_raw_ret - self.beta * self.mix_ret
        print "hedged_ret\t\t{}".format(self.real_ret)  # 每个策略的real收益率
        self.StandProduct.target_real_ret_dict[strategy] = self.real_ret
        self.StandProduct.target_pre_value_dict[strategy] = self.target_pre_value
        print self.StandProduct.target_real_ret_dict



class Unilateral():
    def __init__(self, StandProduct, TotalAlpha):
        self.StandProduct = StandProduct
        self.TotalAlpha = TotalAlpha
        self.calculate_single_side()

    def calculate_single_side(self):
        # 单边
        # print self.single_side
        futpx_yesterday = self.TotalAlpha.futpx_yesterday
        futpx_today = self.TotalAlpha.futpx_today
        m = pd.merge(futpx_today, self.TotalAlpha.single_side, how='right', right_on=['long', 'contract_name'], left_on=['type', 'contract_name'])
        m = pd.merge(futpx_yesterday, m, how='right', right_on=['long', 'contract_name'], left_on=['type', 'contract_name'])
        if not len(m):
            print "No single side holding yesterday"
            self.single_side = 0
            return
        m['earning'] = m.apply(lambda x: x['settle_y'] - x['settle_x'], axis=1)
        # print m
        self.single_side = sum(m['long_amount'] * m['earning'] * m['multiply_x'])
        print "single_side\t\t{}".format(self.single_side)




class CrossIndex():
    def __init__(self, StandProduct, TotalAlpha):
        self.StandProduct = StandProduct
        self.TotalAlpha = TotalAlpha
        self.calculate_cross_index()

    def calculate_cross_index(self):
        # 跨指数收益
        futpx_yesterday = self.TotalAlpha.futpx_yesterday
        futpx_today = self.TotalAlpha.futpx_today

        m = pd.merge(futpx_today, self.TotalAlpha.cross_index, how='right', right_on=['long', 'contract_name'], left_on=['type', 'contract_name'])
        m = pd.merge(futpx_yesterday, m, how='right', right_on=['long', 'contract_name'], left_on=['type', 'contract_name'])
        self.cross_index = sum(m['long_amount'] * (m['target_y'] - m['target_x']) * m['multiply_x'])  # y是今天的，x是昨天的
        print "cross_index\t\t{}".format(self.cross_index)
