import json
import time

import requests
import xlwings as xw
import pandas as pd


def isin_collect(x):
    List_name = []
    List_name.append(['isin code'])
    for item in x:
        item = item.split(',')
        ISIN = item[0]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        List_name.append(tem_list)
    return List_name


def price_collect(x):
    price_name = []
    price_name.append(['آخرین معامله'])
    for item in x:
        item = item.split(',')
        ISIN = item[8]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        price_name.append(tem_list)
    return price_name


def close_collect(x):
    close_name = []
    close_name.append(['قیمت پایانی'])
    for item in x:
        item = item.split(',')
        ISIN = item[11]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        close_name.append(tem_list)
    return close_name


def name_collect(x):
    name_name = []
    name_name.append(['نماد'])
    for item in x:
        item = item.split(',')
        ISIN = item[26]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        name_name.append(tem_list)
    return name_name


def sell_collect(x):
    sell_name = []
    sell_name.append(['بهترین مظنه فروش'])
    for item in x:
        item = item.split(',')
        ISIN = item[5]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        sell_name.append(tem_list)
    return sell_name


def sell_count_collect(x):
    sellc_name = []
    sellc_name.append(['حجم بهترین مظنه فروش'])
    for item in x:
        item = item.split(',')
        ISIN = item[7]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        sellc_name.append(tem_list)
    return sellc_name


def buy_collect(x):
    buy_name = []
    buy_name.append(['بهترین مظنه خرید'])
    for item in x:
        item = item.split(',')
        ISIN = item[4]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        buy_name.append(tem_list)
    return buy_name


def buyc_collect(x):
    buyc_name = []
    buyc_name.append(['حجم بهترین مظنه خرید'])
    for item in x:
        item = item.split(',')
        ISIN = item[6]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        buyc_name.append(tem_list)
    return buyc_name


def val_collect(x):
    val_name = []
    val_name.append(['حجم'])
    for item in x:
        item = item.split(',')
        ISIN = item[14]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        val_name.append(tem_list)
    return val_name

def value_collect(x):
    val_name = []
    val_name.append(['ارزش'])
    for item in x:
        item = item.split(',')
        ISIN = item[16]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        val_name.append(tem_list)
    return val_name

def diff_collect(x):
    diff_name = []
    diff_name.append(['تغییر'])
    for item in x:
        item = item.split(',')
        ISIN = item[9]
        ISIN = ISIN.replace('[', '')
        ISIN = ISIN.replace('"', '')
        tem_list = [ISIN]
        diff_name.append(tem_list)
    return diff_name

while True:
    url = "http://mdapi.tadbirrlc.com/api/Symbol/all"

    response = requests.get(url)
    data = response.text

    parsed = json.loads(data)

    x = parsed['List'].split('],')
    time.sleep(5)

    wb = xw.Book('test.xlsx')
    worksheet = wb.sheets('Sheet1')
    worksheet.range('A1').value = isin_collect(x)
    worksheet.range('B1').value = name_collect(x)
    worksheet.range('C1').value = price_collect(x)
    worksheet.range('D1').value = close_collect(x)
    worksheet.range('E1').value = sell_collect(x)
    worksheet.range('F1').value = sell_count_collect(x)
    worksheet.range('G1').value = buy_collect(x)
    worksheet.range('H1').value = buyc_collect(x)
    worksheet.range('I1').value = val_collect(x)
    worksheet.range('j1').value = value_collect(x)
    worksheet.range('K1').value = diff_collect(x)
