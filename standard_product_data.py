# -*- coding: gbk -*-

from get_trading_data import *
import os
import shutil
import utilities as ut

class StandardProduct:
    def __init__(self, product_id, date=trading_data.date):
        self.t = trading_data
        self.product_id = product_id
        self.path = "{}/accounts-{}/{}".format(root_path, product_id, date.replace('-',''))
        self.copy_accounts_files()
        self.read_csv()
        self.read_trading_record()
        # self.read_wh_target_file()
        self.read_pre_stock_target()
        self.read_pre_future_target()
        self.target_real_ret_dict = {}
        self.target_pre_value_dict = {}

    def copy_accounts_files(self):
        product_id = self.product_id
        from_path = "{}/accounts-{}/{}/".format(root_path, product_id, self.t.yesterday.replace('-', ''))
        to_path = "{}/accounts-{}/{}/".format(root_path, product_id, self.t.date.replace('-', ''))
        file = "accounts.xlsx"
        shutil.copy(from_path + file, (to_path + file).encode('gbk'))  # 把昨天的accounts.xlsx复制到今天的文件夹下

    # 存储六张表
    def read_csv(self):
        self.csv = {}
        for x in ["asset.csv","asset_fut.csv","holding.csv","holding_fut.csv","trade.csv","trade_fut.csv"]:
            full_path = "{}/{}".format(self.path, x)
            if not os.path.exists(full_path):
                raw_input("[Product {}] No such trading file: {}".format(self.product_id, full_path))
            self.csv[x.split('.')[0]] = pd.read_csv(full_path, encoding='gbk')#写到一个dict里面

        self.csv['holding_pre'] = pd.read_csv("{}/accounts-{}/{}/{}".format(root_path, self.product_id, self.t.yesterday.replace('-',''), "holding.csv"), encoding='gbk')
        self.csv['holding_fut_pre'] = pd.read_csv("{}/accounts-{}/{}/{}".format(root_path, self.product_id, self.t.yesterday.replace('-', ''), "holding_fut.csv"),encoding='gbk')
        self.csv['trade']['Commission_type'] = self.csv['trade']['Id'].apply(ut.distinguish_commission_type)
        if self.csv['trade'].empty:
            self.csv['trade']['Commission_cal'] = self.csv['trade']['Id']
        else:
            self.csv['trade']['Commission_cal'] = self.csv['trade'].apply(lambda x: (3.0/10000.0 if x['Direction'] == 'buy' else 13.0/10000.0) * x['Price'] * x['Volume'], axis = 1)
        self.csv['holding']['Commission_type'] = self.csv['holding']['Id'].apply(ut.distinguish_commission_type)
        self.csv['holding_pre']['Commission_type'] = self.csv['holding_pre']['Id'].apply(ut.distinguish_commission_type)
        self.csv['holding_fut']['Contract_name'] = self.csv['holding_fut']['Id'].apply(lambda x: self.month2contract(x[2:]))
        self.csv['holding_fut_pre']['Contract_name'] = self.csv['holding_fut_pre']['Id'].apply(lambda x: self.month2contract(x[2:]))
        self.csv['trade_fut']['Contract_name'] = self.csv['trade_fut']['Id'].apply(lambda x: self.month2contract(x[2:]))

    #读交易记录模板的假设是我们不改变模板的列的原始位置
    def read_trading_record(self):
        self.cross_index = {}
        self.trading_record = {}
        df = pd.read_excel(u"{}/Happy trading day/日交易记录/交易记录模板-NO{}.xlsx".format(root_path_first, self.product_id), skiprows=4, encoding='gbk')
        columns = ['date', 'ipo', 'buy_back', 'transfer', 'suspension', 'x', 's4', 'x', 's6' ,'x', 'event', 'x', 'single_if', 'single_ih', 'single_ic', 'cross_index']
        if len(df.columns) != len(columns):
            df = df.iloc[:, range(len(columns))]
        df.columns = ['date', 'ipo', 'buy_back', 'transfer', 'suspension', 'x', 's4', 'x', 's6', 'x', 'event', 'x', 'single_if', 'single_ih', 'single_ic', 'cross_index']
        df = df.set_index(['date'])
        self.cross_index['today'] = ut.cut_cross_fut(ut.remove_pre_delivery_month(df.loc[self.t.date, 'cross_index'], self.t.delivery_month))
        self.cross_index['yesterday'] = ut.cut_cross_fut(ut.remove_pre_delivery_month(df.loc[self.t.yesterday, 'cross_index'], self.t.delivery_month))
        self.trading_record['s6'] = ut.cut_cross_fut(df.loc[self.t.yesterday, 's6'], prop = 4)

        self.trading_record['s4'] = ut.cut_cross_fut(df.loc[self.t.yesterday, 's4'], prop = 4)
        self.trading_record['suspension'] = ut.cut_cross_fut(df.loc[self.t.yesterday, 'suspension'], prop = 4)
        self.trading_record['event'] = ut.cut_cross_fut(df.loc[self.t.yesterday, 'event'], prop = 4)
        self.trading_record['cross_index'] = ut.cut_cross_fut(ut.remove_pre_delivery_month(df.loc[self.t.yesterday, 'cross_index'], self.t.delivery_month))
        x = self.trading_record['suspension']
        y = self.trading_record['s4']
        z = self.trading_record['s6']

        w = self.trading_record['cross_index']
        u = self.trading_record['event']

        self.index_hedge_target = pd.concat([y, z, u])
        self.beta_index = pd.concat([x, y, z, u])
        self.all_future = pd.concat([x, y, z, w, u])


    def read_wh_target_file(self):
        self.target = {}

        from_path = "//shiming/trading/target_files/{}/".format(self.product_id)
        to_path = u"//shiming/accounts/Happy trading day/target/{}号/{}/".format(self.product_id, self.t.date)
        dirpath,dirnames,filenames = next(os.walk(from_path))


        for x in ["final_stock.csv", "final_future.csv"]:
            if x in filenames:
                self.target[x.split('.')[0]] = pd.read_csv("{}/{}".format(from_path, x)).fillna(0)
                filenames.remove(x)
            else:
                self.target[x.split('.')[0]] = None
                print "[Product {}] No {}.csv today!".format(self.product_id, x)
            # 这里还要对其他csv按照一个future一个stock配对


    # 现在不这么玩了，不用从happy trading day里面读target
    def read_pre_stock_target(self):
        # target_pre_folder = u"{}/Happy trading day/target/{}号/{}".format(root_path, self.product_id, self.t.yesterday.replace('-', ''))
        target_pre_folder = u"{}/python/YZJ/target/{}号/{}".format(root_path_first, self.product_id, self.t.yesterday.replace('-', ''))
        target_pre_paths = os.listdir(target_pre_folder)  #读取上面这个文件夹里的target文件
        self.stock_target_pre = {}
        for target_pre_path in target_pre_paths:
            if target_pre_path.startswith('stock'):
                if (target_pre_path!='stock_final.csv'):
                    strategy = ut.extract_strategy_type(target_pre_path) #先判断path里的文件属于哪个策略，s4，s6？
                    df = pd.read_csv(u"//shiming/accounts/python/YZJ/target/{}号/{}/{}".format(self.product_id, self.t.yesterday.replace('-', ''), target_pre_path))
                    self.stock_target_pre[strategy[0]] = df.loc[:,['coid','stockhand']].fillna(0)
            elif target_pre_path == 'final_stock.csv':
                df = pd.read_csv(u"//shiming/accounts/python/YZJ/target/{}号/{}/{}".format(self.product_id, self.t.yesterday.replace('-', ''), target_pre_path))
                self.stock_target_pre['final'] = df.loc[:, ['coid', 'stockhand']].fillna(0)


    def read_pre_future_target(self):
        target_pre_folder = u"{}/python/YZJ/target/{}号/{}".format(root_path_first, self.product_id,self.t.yesterday.replace('-', ''))
        target_pre_paths = os.listdir(target_pre_folder)  # 读取上面这个文件夹里的target文件
        self.future_target_pre = {}
        for target_pre_path in target_pre_paths:
            if target_pre_path.startswith('future'):
                if target_pre_path != 'future_final.csv':
                    strategy = ut.extract_strategy_type(target_pre_path)  # 先判断path里的文件属于哪个策略，s4，s6？
                    df = pd.read_csv(u"//shiming/accounts/python/YZJ/target/{}号/{}/{}".format(self.product_id, self.t.yesterday.replace('-', ''), target_pre_path))
                    self.future_target_pre[strategy[0]] = df.loc[:, ['future_number_300', 'future_number_500', 'future_number_50']].fillna(0)
            elif target_pre_path == 'final_future.csv':
                df = pd.read_csv(u"//shiming/accounts/python/YZJ/target/{}号/{}/{}".format(self.product_id, self.t.yesterday.replace('-', ''), target_pre_path))
                self.future_target_pre['final'] = df.loc[:, ['future_number_300', 'future_number_500', 'future_number_50']].fillna(0)

    def month2contract(self, month, today=True):#month='1605'
        dm = self.t.delivery_month
        if today:
            if month in dm['delivery'].values:
                contract = dm[dm['delivery']==month].iloc[0,0]
                return contract
            else:
                print "*****Future contract {} not exist in product {}!!!*****".format(month, self.product_id)

        else:
            if month in dm['pre_delivery'].values:
                contract = dm[dm['pre_delivery']==month].iloc[0,0]
                return contract
            else:
                print "*****Future contract {} not exist in product {}!!!*****".format(month, self.product_id)

