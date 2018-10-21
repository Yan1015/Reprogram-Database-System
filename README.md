# Reprogram-Database-System
## Introduction
This project is to reprogram the database system of a Chinese Investment Management Company based on Python. 

In the company, there are many products which apply different trading strategy. For the trading group, they will adjust the trading strategy based on everyday's perfromance, so we need to record all the datails of the performance of all the products.

The database system aims to monitor the difference between the predictive results of trading strategies and the actual results of trading operations. Commonly they should have slight difference, so this system can help to monitor problems which happen during the daily trading operations.

The company is becoming bigger and bigger, and form the beginning the target market is only stock market, now there are many strategies targeting to stock, future and options market. So we need to reprogram the database system to make it adaptable for new products.

## Code Structure
accounts__main__.py is the main function of all the system.

accounts_info.py is to get all the information of each product's account.

get_trading_data.py helps us to get the daily trading data, such as stock price, future price and option price.

standard_product_data.py is to make every product as a class.

ret_decomposition.py calculates all the details, such as the profits, costs and alpha return.

utilities.py contains all the other utilitiy function which we would use.




