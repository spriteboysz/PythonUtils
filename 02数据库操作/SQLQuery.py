#! /usr/bin/env python
# coding=utf-8
"""
Author: Deean
Date: 2023-12-09 23:15
FileName: 02数据库操作
Description:SQLQuery.py 
"""

import pandas as pd
from sqlalchemy import create_engine

# 定义数据库连接参数
__HOST__ = 'localhost'
__PORT__ = '3306'
__USER__ = 'root'
__PASSWORD__ = 'www.huawei.com'


def get_engine(db='P1757'):
    url = 'mysql+pymysql://{}:{}@{}:{}/{}?charset=utf8'.format(
        __USER__, __PASSWORD__, __HOST__, __PORT__, db)
    return create_engine(url, echo=True)


def get_sql_table(db='bookmall', table='users'):
    engine = get_engine(db)
    with engine.connect() as conn:
        sql = 'select * from ' + table
        print(pd.read_sql(sql, conn))
        return pd.read_sql_table(table, conn)


if __name__ == '__main__':
    df = get_sql_table('bookmall', 'users')
    print(df)
