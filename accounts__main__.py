# -*- coding: gbk -*-
from OutPut import *
import datetime as dt
import utilities as ut
# import pandas as pd
# from standard_product_data import *
# from accounts_info import *
# from ret_decomposition import *
# import openpyxl as ox

# download_xuntou()
# download_ctp()

def back_up_accounts_raw_data_hao(id, file_list):
    from_path = r"//shiming/accounts/" + str(id) + 'hao/'
    date = dt.datetime.today().date()
    to_path = ur'//SHIMING/accounts/python/back_up/accounts_raw_data备份/{}/{}'.format(str(id), str(date).replace('-', ''))
    if not os.path.exists(to_path):
        os.mkdir(to_path)
    for file in file_list:
        if not os.path.exists(to_path + '/' + file):
            shutil.copy(from_path + file, (to_path + '/' + file).encode('gbk'))

def back_up_accounts_raw_data_CTP(id, file_list):
    from_path = r"//shiming/accounts/CSVDownloader/Download/CTP{}/".format(str(id))
    date = dt.datetime.today().date()
    to_path = ur'//SHIMING/accounts/python/back_up/accounts_raw_data备份/{}/{}'.format(str(id), str(date).replace('-', ''))
    if not os.path.exists(to_path):
        os.mkdir(to_path)
    for file in file_list:
        if not os.path.exists(to_path + '/' + file):
            shutil.copy(from_path + file, (to_path + '/' + file).encode('gbk'))

def back_up_accounts_raw_data_XT(id, file_list):
    from_path = r"//shiming/accounts/CSVDownloader/Download/XT{}/".format(str(id))
    date = dt.datetime.today().date()
    to_path = ur'//SHIMING/accounts/python/back_up/accounts_raw_data备份/{}/{}'.format(str(id), str(date).replace('-', ''))
    if not os.path.exists(to_path):
        os.mkdir(to_path)
    for file in file_list:
        if not os.path.exists(to_path + '/' + file):
            shutil.copy(from_path + file, (to_path + '/' + file).encode('gbk'))

def back_up_accounts_raw_data():
    # 4, 10, 11, 19, 101, 102, 107, 203, 208, 301, 309, 312
    back_up_accounts_raw_data_hao(1, ['holding.xls', 'trade.xls'])
    back_up_accounts_raw_data_hao(2, ['asset.csv', 'holding.xls', 'trade.xls', 'asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'])
    # back_up_accounts_raw_data_hao(4, ['asset_fut.txt', 'holding_fut.csv', 'trade_fut.csv', 'asset.csv', 'holding.xls', 'trade.xls'])
    back_up_accounts_raw_data_hao(5, ['asset_fut.txt', 'holding_fut.csv', 'trade_fut.csv'])
    back_up_accounts_raw_data_hao(7, ['asset_fut.txt', 'holding.xls', 'trade.xls', 'holding_fut.csv', 'trade_fut.csv'])
    back_up_accounts_raw_data_hao(8, ['asset_fut.txt', 'holding.xls', 'holding_fut.csv', 'trade.xls', 'trade_fut.csv'])
    back_up_accounts_raw_data_hao(12, ['asset_fut.txt', 'holding.xls', 'holding_fut_d.csv', 'trade.xls', 'trade_fut_d.csv', 'asset.csv'])
    back_up_accounts_raw_data_hao(15, ['asset_fut.csv', 'holding.xls', 'holding_fut.xls', 'trade.xls', 'trade_fut.xls'])
    back_up_accounts_raw_data_hao(16, ['asset_fut.txt', 'holding.xls', 'trade.xls', 'holding_fut.csv', 'trade_fut.csv'])
    back_up_accounts_raw_data_hao(17, ['asset_fut.txt', 'holding.xls', 'trade.xls', 'holding_fut.csv', 'trade_fut.csv', 'holding_l.xls', 'trade_l.xls'])
    back_up_accounts_raw_data_CTP(1, ['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'])
    back_up_accounts_raw_data_CTP(2, ['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'])
    back_up_accounts_raw_data_CTP(12, ['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'])
    # back_up_accounts_raw_data_CTP(13, ['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'])
    # back_up_accounts_raw_data_XT(0, ['asset.csv', 'holding.csv', 'trade.csv', 'asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'])

def update_chenjie_holding():
    for id in [1, 2, 4, 5, 7, 8, 10, 11, 16, 17, 19, 101, 102, 107, 203, 208, 301, 309, 311, 312, 313]:
        date = dt.datetime.today().date()
        from_path = r"//shiming/accounts/python/YZJ/accounts/accounts-{}/{}/".format(str(id), str(date).replace('-', ''))
        to_path = ur"//shiming/accounts/python/holdings/{}/{}".format(str(id), str(date).replace('-', ''))
        if not os.path.exists(to_path):
            os.mkdir(to_path)
        for file_name in ['holding.csv']:
            if not os.path.exists(to_path + '/' + file_name):
                shutil.copy(from_path + file_name, (to_path + '/' + file_name).encode('gbk'))

    # for id in [8]:
    #     date = dt.datetime.today().date()
    #     from_path = r"//shiming/accounts/{}hao/".format(str(id), str(date).replace('-', ''))
    #     to_path = ur"//shiming/accounts/python/holdings/{}/{}".format(str(id), str(date).replace('-', ''))
    #     if not os.path.exists(to_path):
    #         os.mkdir(to_path)
    #     for file_name in ['holding.xls']:
    #         if not os.path.exists(to_path + '/' + file_name):
    #             shutil.copy(from_path + file_name, (to_path + '/' + file_name).encode('gbk'))






def __main__():

    print u'-----------------------1号raw_data Transform 开始-----------------------'
    e1 = Transform(1)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e1.holding_zhongxin_security_1()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e1.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费",  u"买卖"])
    print u"-----------------------现在正在转换三个fut文件,如报错请检查该文件-----------------------"
    e1.T(['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'], dir_path="//shiming/accounts/CSVDownloader/Download/CTP1")
    print u'-----------------------1号raw_data Transform 结束-----------------------'

    print u'-----------------------2号raw_data Transform 开始-----------------------'
    e2 = Transform(2)
    print u"-----------------------现在正在转换asset,如报错请检查该文件-----------------------"
    e2.T('asset.csv')
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e2.T('trade.xls', '\t', [u"证券代码", u"成交均价", u"成交数量", u"手续费",  u"操作"])
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e2.T('holding.xls', '\t', [u"证券代码", u"股票余额", u"市价",  u"成本价", u"盈亏"])
    e2.transform_cost('holding.csv')
    print u"-----------------------现在正在转换三个fut文件,如报错请检查该文件-----------------------"
    e2.T(['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'], dir_path="//shiming/accounts/CSVDownloader/Download/CTP2")
    print u'-----------------------2号raw_data Transform 结束-----------------------'

    print u'-----------------------4号raw_data Transform 开始-----------------------'
    e4 = Transform(4)
    print u"-----------------------现在正在转换asset,如报错请检查该文件-----------------------"
    e4.T('asset.csv')
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e4.T('trade.xls', '\t', [u"证券代码", u"成交均价", u"成交数量", u"手续费", u"操作"])
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e4.T('holding.xls', '\t', [u"证券代码", u"股票余额", u"市价", u"成本价", u"盈亏"])
    e4.transform_cost('holding.csv')
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e4.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e4.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e4.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------4号raw_data Transform 结束-----------------------'

    print u'-----------------------5号raw_data Transform 开始-----------------------'
    e5 = Transform(5)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e5.holding_zhongxin_security_17()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e5.transform_17('trade.xls', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"方向"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e5.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e5.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e5.transform_asset_fut17([u"动态权益", u"当前保证金", u"可用资金", u"多单总市值", u"空单总市值"])
    print u'-----------------------5号raw_data Transform 结束-----------------------'

    print u'-----------------------7号raw_data Transform 开始-----------------------'
    e7 = Transform(7)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e7.holding_zhongxin_security_normal()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e7.transform_normal('trade.xls',[u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e7.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e7.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e7.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------7号raw_data Transform 结束-----------------------'

    print u'-----------------------8号raw_data Transform 开始-----------------------'
    e8 = Transform(8)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e8.holding_zhongxin_security_normal()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e8.transform_normal('trade.xls',[u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e8.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e8.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e8.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------8号raw_data Transform 结束-----------------------'

    print u'-----------------------10号raw_data Transform 开始-----------------------'
    e10 = Transform(10)
    print u"-----------------------现在正在转换asset,如报错请检查该文件-----------------------"
    e10.T('asset.csv')
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e10.T('trade.xls', '\t', [u"证券代码", u"成交均价", u"成交数量", u"手续费", u"操作"])
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e10.T('holding.xls', '\t', [u"证券代码", u"股票余额", u"市价", u"成本价", u"盈亏"])
    e10.transform_cost('holding.csv')
    print u'-----------------------10号raw_data Transform 结束-----------------------'

    print u'-----------------------11号raw_data Transform 开始-----------------------'
    e11 = Transform(11)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e11.holding_zhongxin_security_normal()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e11.transform_normal('trade.xls',[u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e11.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e11.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e11.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------11号raw_data Transform 结束-----------------------'

    # print u'-----------------------12号raw_data Transform 开始-----------------------'
    # e12 = Transform(12)
    # print u"-----------------------现在正在转换asset,如报错请检查该文件-----------------------"
    # e12.T('asset.csv')
    # print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    # e12.T('trade.xls', '\t', [u"证券代码", u"成交均价", u"成交数量", u"手续费",  u"操作"])
    # print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    # e12.T('holding.xls', '\t', [u"证券代码", u"股票余额", u"市价", u"成本价", u"盈亏"])
    # e12.transform_cost('holding.csv')
    # print u"-----------------------现在正在转换三个fut文件,如报错请检查该文件-----------------------"
    # e12.T(['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'], dir_path="//shiming/accounts/CSVDownloader/Download/CTP12")
    # print u'-----------------------12号raw_data Transform 结束-----------------------'

    # print u'-----------------------13号raw_data Transform 开始-----------------------'
    # e13 = Transform(13)
    # print u"-----------------------现在正在转换asset,如报错请检查该文件-----------------------"
    # e13.T('asset.csv', ',', [u"冻结金额", u"可用金额", u"总市值", u"总资产"])
    # print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    # e13.T('trade.csv', ',', [u"证券代码", u"成交价格", u"成交数量", u"手续费",  u"操作"])
    # print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    # e13.T('holding.csv', ',', [u"证券代码", u"当前拥股", u"最新价", u"成本价", u"盈亏"])
    # e13.transform_cost('holding.csv')
    # print u"-----------------------现在正在转换三个fut文件,如报错请检查该文件-----------------------"
    # e13.T(['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'], dir_path="//shiming/accounts/CSVDownloader/Download/CTP13")
    # print u'-----------------------13号raw_data Transform 结束-----------------------'

    print u'-----------------------16号raw_data Transform 开始-----------------------'
    e16 = Transform(16)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e16.holding_zhongxin_security_normal()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e16.transform_normal('trade.xls',[u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e16.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e16.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e16.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------16号raw_data Transform 结束-----------------------'

    print u'-----------------------17号raw_data Transform 开始-----------------------'
    e17 = Transform(17)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e17.holding_zhongxin_security_17()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e17.transform_17('trade.xls', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"方向"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e17.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e17.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e17.transform_asset_fut17([u"动态权益", u"当前保证金", u"可用资金", u"多单总市值", u"空单总市值"])

    print u"-----------------------现在正在转换holding_l,如报错请检查该文件-----------------------"
    e17.holding_zhongxin_security_17_low()
    print u"-----------------------现在正在转换trade_l,如报错请检查该文件-----------------------"
    e17.transform_trade_17()

    print u"-----------------------现在正在合并17号上下层-----------------------"
    e17.plus_holding_17hao()
    e17.concat_low_to_high(0, 'trade.csv', 'trade_l.csv', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"],
                           ['Id', 'Price', 'Volume', 'Commission', 'Direction'],
                           dir_path="//shiming/accounts/17hao/transform",
                           transform_path="//shiming/accounts/17hao/transform")
    e17.plus_low_to_high('asset.csv', 'asset_l.csv', ['Frozen', 'Avaliable', 'MarketCap', 'TotalAsset'],
                         ['Frozen', 'Avaliable', 'MarketCap', 'TotalAsset'],
                         dir_path="//shiming/accounts/17hao/transform",
                         transform_path="//shiming/accounts/17hao/transform")
    print u'-----------------------17号raw_data Transform 结束-----------------------'

    print u'-----------------------19号raw_data Transform 开始-----------------------'
    e19 = Transform(19)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e19.holding_zhongxin_security_1()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e19.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费",  u"买卖"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e19.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e19.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e19.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------19号raw_data Transform 结束-----------------------'

    print u'-----------------------101号raw_data Transform 开始-----------------------'
    e101 = Transform(101)
    print u"-----------------------现在正在转换asset,如报错请检查该文件-----------------------"
    e101.T('asset.csv')
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e101.T('trade.xls', '\t', [u"证券代码", u"成交均价", u"成交数量", u"手续费",  u"操作"])
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e101.T('holding.xls', '\t', [u"证券代码", u"股票余额", u"市价",  u"成本价", u"盈亏"])
    e101.transform_cost('holding.csv')
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e101.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e101.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e101.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------101号raw_data Transform 结束-----------------------'

    print u'-----------------------102号raw_data Transform 开始-----------------------'
    e102 = Transform(102)
    print u"-----------------------现在正在转换asset,如报错请检查该文件-----------------------"
    e102.T('asset.csv')
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e102.T('trade.xls', '\t', [u"证券代码", u"成交均价", u"成交数量", u"手续费",  u"操作"])
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e102.T('holding.xls', '\t', [u"证券代码", u"股票余额", u"市价",  u"成本价", u"盈亏"])
    e102.transform_cost('holding.csv')
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e102.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e102.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e102.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------102号raw_data Transform 结束-----------------------'

    print u'-----------------------107号raw_data Transform 开始-----------------------'
    e107 = Transform(107)
    print u"-----------------------现在正在转换asset,如报错请检查该文件-----------------------"
    e107.T('asset.csv')
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e107.T('trade.xls', '\t', [u"证券代码", u"成交均价", u"成交数量", u"手续费",  u"操作"])
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e107.T('holding.xls', '\t', [u"证券代码", u"股票余额", u"市价",  u"成本价", u"盈亏"])
    e107.transform_cost('holding.csv')
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e107.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e107.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e107.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------107号raw_data Transform 结束-----------------------'

    print u'-----------------------203号raw_data Transform 开始-----------------------'
    e203 = Transform(203)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e203.holding_zhongxin_security_normal(input_file='holding.xls', sep='\t', asset_columns=[u'冻结金额', u'可用', u'参考市值', u'资产'], holding_columns=[u'证券代码', u'证券数量', u'当前价', u'参考成本价', u'参考盈亏'])
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e203.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e203.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    e203.transform_holding_fut('holding_fut_l.csv', [u'合约', u'买卖', u'总持仓'], 1)
    e203.concat_low_to_high(1, 'holding_fut.csv', 'holding_fut_l.csv', ['Id', 'Direction', 'Holding'], ['Id', 'Direction', 'Holding'], dir_path="//shiming/accounts/203hao/transform", transform_path = "//SHIMING/accounts/203hao/transform")
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e203.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    e203.transform_trade_fut('trade_fut_l.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'],1)
    e203.concat_low_to_high(1, 'trade_fut.csv', 'trade_fut_l.csv', ['Id', 'Direction', 'Commission', 'Price', 'Volume'], ['Id', 'Direction', 'Commission', 'Price', 'Volume'], dir_path="//shiming/accounts/203hao/transform", transform_path = "//SHIMING/accounts/203hao/transform")
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e203.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    e203.transform_asset_fut_normal_l(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    e203.plus_low_to_high('asset_fut.csv', 'asset_fut_l.csv', ['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'], ['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'], dir_path="//SHIMING/accounts/203hao/transform", transform_path = "//SHIMING/accounts/203hao/transform")
    print u'-----------------------203号raw_data Transform 结束-----------------------'

    print u'-----------------------208号raw_data Transform 开始-----------------------'
    e208 = Transform(208)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e208.holding_zhongxin_security_normal(input_file='holding.xls', sep='\t', asset_columns=[u'冻结金额', u'可用', u'参考市值', u'资产'], holding_columns=[u'证券代码', u'库存数量', u'当前价', u'参考成本价', u'参考盈亏'])
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e208.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u"-----------------------现在正在转换holding_fut,如报错请检查该文件-----------------------"
    e208.transform_holding_fut('holding_fut.csv', [u'合约', u'买卖', u'总持仓'])
    print u"-----------------------现在正在转换trade_fut,如报错请检查该文件-----------------------"
    e208.transform_trade_fut('trade_fut.csv', [u'合约', u'买卖', u'手续费', u'成交价格', u'成交手数'])
    print u"-----------------------现在正在转换asset_fut,如报错请检查该文件-----------------------"
    e208.transform_asset_fut_normal(['Balance', 'CurrMargin', 'Avaliable', 'Long_MarketCap', 'Short_MarketCap'])
    print u'-----------------------208号raw_data Transform 结束-----------------------'

    print u'-----------------------301号raw_data Transform 开始-----------------------'
    e301 = Transform(301)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e301.holding_zhongxin_security_normal(input_file='holding.xls', sep='\t', asset_columns=[u'冻结金额', u'可用', u'参考市值', u'资产'], holding_columns=[u'证券代码', u'参考持股', u'当前价', u'成本价', u'浮动盈亏'])
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e301.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费",  u"买卖"])
    print u'-----------------------301号raw_data Transform 结束-----------------------'

    print u'-----------------------309号raw_data Transform 开始-----------------------'
    e309 = Transform(309)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e309.holding_zhongxin_security_normal()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e309.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费",  u"买卖标志"])
    print u'-----------------------309号raw_data Transform 结束-----------------------'

    print u'-----------------------311号raw_data Transform 开始-----------------------'
    e311 = Transform(311)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e311.holding_zhongxin_security_normal()
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e311.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u'-----------------------311号raw_data Transform 结束-----------------------'

    print u'-----------------------312号raw_data Transform 开始-----------------------'
    e312 = Transform(312)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e312.holding_zhongxin_security_normal(input_file='holding.xls', sep='\t', asset_columns=[u'冻结金额', u'可用', u'参考市值', u'资产'], holding_columns=[u'证券代码', u'证券数量', u'当前价', u'参考成本价', u'参考盈亏'])
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e312.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u'-----------------------312号raw_data Transform 结束-----------------------'

    print u'-----------------------313号raw_data Transform 开始-----------------------'
    e313 = Transform(313)
    print u"-----------------------现在正在转换holding,如报错请检查该文件-----------------------"
    e313.holding_zhongxin_security_normal(input_file='holding.xls', sep='\t', asset_columns=[u'冻结金额', u'可用', u'参考市值', u'资产'], holding_columns=[u'证券代码', u'证券数量', u'参考市价', u'参考成本价', u'参考盈亏'])
    print u"-----------------------现在正在转换trade,如报错请检查该文件-----------------------"
    e313.T('trade.xls', '\t', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标志"])
    print u'-----------------------313号raw_data Transform 结束-----------------------'


    # no more
    # e7 = Transform(7)
    # e7.plus_holding_7hao()
    # e7.concat_low_to_high(0, 'trade.csv', 'trade_l.csv', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"操作"], [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"操作"], dir_path="//shiming/accounts/CSVDownloader/Download/XT7")
    # e7.concat_low_to_high(1, 'holding_fut.csv', 'holding_fut_l.csv', [u"合约代码", u"买卖", u"持仓量"], [u"合约代码", u"多空", u"持仓量"], dir_path="//shiming/accounts/CSVDownloader/Download/XT7")
    # e7.concat_low_to_high(1, 'trade_fut.csv', 'trade_fut_l.csv', [u"合约代码", u"多空", u"手续费", u"成交均价", u"成交量"], [u"合约代码", u"多空", u"手续费", u"成交均价", u"成交量"], dir_path="//shiming/accounts/CSVDownloader/Download/XT7")
    # e7.plus_low_to_high('asset.csv', 'asset_l.csv', [u"冻结金额", u"可用金额", u"总市值", u"总资产"], [u"冻结金额", u"可用金额", u"总市值", u"总资产"], dir_path="//shiming/accounts/CSVDownloader/Download/XT7")
    # e7.plus_low_to_high('asset_fut.csv', 'asset_fut_l.csv', [u"动态权益", u"当前保证金", u"可用资金", u"多单总市值", u"空单总市值"], [u"动态权益", u"当前保证金", u"可用资金", u"多单总市值", u"空单总市值"], dir_path="//shiming/accounts/CSVDownloader/Download/XT7")
    #
    # e11 = Transform(11)
    # e11.holding_zhongxin_security_11()
    # e11.transform_trade_11()
    # e11.transform_asset_fut11([u"动态权益", u"当前保证金", u"可用资金", u"多单总市值", u"空单总市值"])
    # e11.plus_holding_11hao()
    # e11.concat_low_to_high(0, 'trade.csv', 'trade_l.csv', [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖"], [u"证券代码", u"成交价格", u"成交数量", u"手续费", u"买卖标记"], dir_path="//shiming/accounts/CSVDownloader/Download/XT11")
    # e11.concat_low_to_high(1, 'holding_fut.csv', 'holding_fut_l.csv', [u"合约", u"买卖", u"总持仓"], [u"合约代码", u"买卖", u"持仓量"], dir_path="//shiming/accounts/CSVDownloader/Download/XT11")
    # e11.concat_low_to_high(1, 'trade_fut.csv', 'trade_fut_l.csv', [u"合约", u"买卖", u"手续费", u"成交价格", u"成交手数"], [u"合约代码", u"多空", u"手续费", u"成交均价", u"成交量"], dir_path="//shiming/accounts/CSVDownloader/Download/XT11")
    # e11.plus_low_to_high('asset.csv', 'asset_l.csv', [u"冻结金额", u"可用金额", u"总市值", u"总资产"], [u"冻结金额", u"可用金额", u"总市值", u"总资产"], dir_path="//shiming/accounts/CSVDownloader/Download/XT11")
    # e11.plus_low_to_high('asset_fut.csv', 'asset_fut_l.csv', [u"动态权益", u"当前保证金", u"可用资金", u"多单总市值", u"空单总市值"], [u"动态权益", u"当前保证金", u"可用资金", u"多单总市值", u"空单总市值"], dir_path="//shiming/accounts/CSVDownloader/Download/XT11")

    print u'-----------------------Transform 数据转移开始-----------------------'
    # for pid in [0]:
    #     # 0
    #     e = Transform(pid)
    #     e.T(['asset.csv', 'trade.csv', 'holding.csv', 'asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'], dir_path="//shiming/accounts/CSVDownloader/Download/XT{}".format(pid))
    for pid in [5, 7, 11, 8, 16, 17]:
        # 5, 7, 11, 8, 16, 17
        e = Transform(pid)
        e.T(['asset.csv', 'trade.csv', 'holding.csv', 'asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'], dir_path="//shiming/accounts/{}hao/transform".format(pid), process=2)
    for pid in [4, 19, 101, 102, 107]:
        # 4, 19, 101, 102, 107
        e = Transform(pid)
        e.T(['asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'], dir_path="//shiming/accounts/{}hao/transform".format(pid), process=2)
    for pid in [203, 208]:
        # 203, 208
        e = Transform(pid)
        e.T(['asset.csv', 'holding.csv', 'asset_fut.csv', 'holding_fut.csv', 'trade_fut.csv'], dir_path="//shiming/accounts/{}hao/transform".format(pid), process=2)
    for pid in [301, 309, 311, 312, 313]:
        # 301, 309, 311, 312, 313
        e = Transform(pid)
        e.T(['asset.csv', 'holding.csv'], dir_path="//shiming/accounts/{}hao/transform".format(pid), process=2)
    print u'-----------------------Transform 数据转移结束-----------------------'

# 以下是对下载的原始数据的备份,要放在__main__()之前

# back_up_accounts_raw_data()
# __main__()
# update_chenjie_holding()


for i in [2]:
    # 1, 2, 4, 5, 7, 11, 8, 12, 16, 17, 19, 101, 102, 107, 208
    print u'目前更新的产品id为：', i
    print u'是否继续,继续输入Y，停止运行输入N'
    ch = raw_input()
    if (ch == 'N') or (ch == 'n'):
        exit()
    standard_product_data = StandardProduct(i)
    asset = Asset(standard_product_data)
    totalalpha = TotalAlpha(standard_product_data)
    alpha = Alpha(standard_product_data, totalalpha)
    unilateral = Unilateral(standard_product_data, totalalpha)
    crossindex = CrossIndex(standard_product_data, totalalpha)
    x = output_accounts(asset, totalalpha)

    print u'-----------------------{}号输出sheet_accounts开始-----------------------'.format(str(i))
    x.output_sheet_accounts()
    print u'-----------------------{}号输出sheet_accounts结束-----------------------'.format(str(i))
    y = output_ret_decomposition(totalalpha, alpha, unilateral, crossindex, asset)
    print u'-----------------------{}号输出sheet_ret_dc开始-----------------------'.format(str(i))
    y.output_sheet_ret()
    print u'-----------------------{}号输出sheet_ret_dc结束-----------------------'.format(str(i))
    z = update(asset)
    if i in [0, 1, 2, 4, 7, 17, 19, 101, 102, 107]:
        # 0, 1, 2, 4, 7, 17, 19, 101, 102
        print u'-----------------------{}号更新summary开始-----------------------'.format(str(i))
        z.update_summary_1()
        print u'-----------------------{}号更新summary结束-----------------------'.format(str(i))
    # if i in [11]:
    #     print u'-----------------------{}号更新summary开始-----------------------'.format(str(i))
    #     z.update_summary_2()
    #     print u'-----------------------{}号更新summary结束-----------------------'.format(str(i))
    print u'-----------------------{}号更新assessreport开始-----------------------'.format(str(i))
    z.update_assessreport()
    print u'-----------------------{}号更新assessreport结束-----------------------'.format(str(i))
    ut.copy_to_account(standard_product_data.product_id, standard_product_data.t.date.replace('-', ''))
