# -*- coding: utf-8 -*-
import click
import pandas as pd
import sys
import openpyxl
from datetime import datetime, timedelta
from openpyxl.chart import (
    LineChart,
    Reference,
)
from openpyxl.chart.axis import DateAxis

INIT_TOTAL_VALUE = 2000000.0
PER_STOCK_BUY_IN_VALUE = 50000.0
MAX_HOLDING_STOCKS_NUM = 30  # 最多30只自选股
DATE_COLUMN = "A"
TOTAL_VALUE_COLUMN = "B"
TOTAL_PROFIT_RATE_COLUMN = "C"
STOCK_CODE_COLUMN = "D"
STOCK_NAME_COLUMN = "E"
HOLDING_VALUE_COLUMN = "F"
OPENING_PRICE_COLUMN = "G"
CLOSING_PRICE_COLUMN = "H"
BUY_IN_PRICE_COLUMN = "I"
PROFIT_COLUMN = "J"
PROFIT_RATE_COLUMN = "K"
PROFIT_HISTORY_SHEET = "ProfitHistory"


class StockInfo(object):
    __slots__ = ["code", "name"]

    def __init__(self, code, name):
        self.code = code
        self.name = name

    def __str__(self):
        return "code: %s, name: %s" % (self.code, self.name)


class DailyStockPriceInfo(object):
    __slots__ = ["date", "stock_code", "opening_price", "closing_price"]

    def __init__(self, date, stock_code, opening_price, closing_price):
        self.date = date
        self.stock_code = stock_code
        self.opening_price = opening_price
        self.closing_price = closing_price


class HoldingStockInfo(object):
    __slots__ = ["stock", "buy_in_value", "buy_in_date", "buy_in_price", "current_value", "current_date", "current_price",
                 "sell_out_date", "sell_out_price", "profit_rate"]

    def __init__(self, stock, buy_in_date, buy_in_price, buy_in_value=PER_STOCK_BUY_IN_VALUE):
        self.stock = stock
        self.buy_in_date = buy_in_date
        self.buy_in_price = buy_in_price
        self.current_value = self.buy_in_value = buy_in_value
        self.current_date = buy_in_date

    def calculate_profit(self, cur_date, cur_price):
        self.current_date = cur_date
        self.current_price = cur_price
        self.current_value = self.buy_in_value * cur_price / self.buy_in_price
        self.profit_rate = (self.current_value -
                            self.buy_in_value)*1.0 / self.buy_in_value

    def sell_out(self, sell_out_date, sell_out_price):
        self.sell_out_date = sell_out_date
        self.sell_out_price = sell_out_price
        self.calculate_profit(sell_out_date, sell_out_price)

    def __str__(self):
        return "stock: %s, buy_in_date: %s, buy_in_price: %.2f, buy_in_value: %s, current_date: %s, current_price: %.2f, current_value: %s, profit_rate: %.2f" % \
            (self.stock, self.buy_in_date, self.buy_in_price,
             self.buy_in_value, self.current_date, self.current_price, self.current_value, 100.0*self.profit_rate) + "%"


class StockMarketHelper(object):
    @staticmethod
    def get_stock_price_info(stock_code, date):
        date = datetime.strptime(date, "%Y/%m/%d").strftime("%Y%m%d")
        url_mode = "http://quotes.money.163.com/service/chddata.html?code=%s&start=%s&end=%s"
        for code_prefix in range(0, 2):
            url = url_mode % ("%d%s" % (
                code_prefix, stock_code), date, date)
            df = pd.read_csv(url, encoding="gb2312")
            if len(df) > 0:
                return DailyStockPriceInfo(date, stock_code, float(df["开盘价"][0]), float(df["收盘价"][0]))
        return None

    @staticmethod
    def is_valid_stock_code(stock_code):
        try:
            int(stock_code)
            return True
        except:
            return False


class InvestmentInfo(object):
    __slots__ = ["holding_stocks", "total_value", "init_value",
                 "cash_value", "stock_value", "profit_rate", "profit_history"]

    def __init__(self, init_value):
        self.stock_value = 0
        self.cash_value = init_value
        self.init_value = init_value
        self.holding_stocks = dict()
        self.profit_history = []

    def __str__(self):
        return "total_value: %.2f, cash_value: %.2f, stock_value: %.2f, profit_rate: %.2f" % \
            (self.total_value, self.cash_value,
             self.stock_value, self.profit_rate*100.0) + "%"

    def is_holding_stock(self, stock_code):
        return stock_code in self.holding_stocks

    def get_holding_stock(self, stock_code):
        return self.holding_stocks.get(stock_code)

    def get_holding_stocks(self):
        return [st_code for st_code in self.holding_stocks]

    def calculate_profit(self, date):
        self.stock_value = 0
        for code, stock_info in self.holding_stocks.items():
            # stock_info.calculate_profit(date)
            self.stock_value += stock_info.current_value
        self.total_value = self.cash_value + self.stock_value
        self.profit_rate = (self.total_value -
                            self.init_value) / self.init_value
        self.profit_history.append((date, self.total_value, self.profit_rate))

    def buy_in_stock(self, stock, date, buy_in_price):
        if stock.code in self.holding_stocks:
            raise RuntimeError(
                "Duplicated holding stock, code %s, name: %s" % (stock.code, stock.name))
        if self.cash_value < PER_STOCK_BUY_IN_VALUE:
            raise RuntimeError(
                "Failed to buy in stock since cash is not enough, current cash: %.2f" % self.cash_value)
        holding_info = HoldingStockInfo(stock, date, buy_in_price)
        self.holding_stocks[stock.code] = holding_info
        self.cash_value -= PER_STOCK_BUY_IN_VALUE
        print("BUY-IN, stock %s, date %s, buy_in_price %.2f, cash_value %.2f" %
              (stock, date, buy_in_price, self.cash_value))

    def sell_out_stock(self, stock_code, sell_out_date):
        if stock_code not in self.holding_stocks:
            raise RuntimeError(
                "Failed to sell out stock %s, because stock does not exit in holding stocks" % (stock_code))
        price_info = StockMarketHelper.get_stock_price_info(
            stock_code, sell_out_date)
        if price_info is None:
            raise RuntimeError("Failed to sell out stock %s, failed to get stock price info of date %s" % (
                stock_code, sell_out_date))
        holding_info = self.holding_stocks.get(stock_code)
        # 当天的开盘价卖出
        holding_info.sell_out(sell_out_date, price_info.opening_price)
        # 收益落袋
        self.cash_value += holding_info.current_value
        print("SELL-OUT, stock info: %s, date: %s, cash_value: %.2f" %
              (holding_info, sell_out_date, self.cash_value))
        # 删除stock
        del self.holding_stocks[stock_code]


def normalize_date(date):
    if isinstance(date, datetime):
        return date.strftime("%Y/%m/%d")
    elif isinstance(date, str):
        return datetime.strptime(date, "%Y/%m/%d").strftime("%Y/%m/%d")


def get_row_range(sheet, date, start_row_hint):
    start_row = -1
    end_row = -1
    # MAX_HOLDING_STOCKS_NUM 用来加快结束查找
    for idx in range(start_row_hint, min(start_row_hint + MAX_HOLDING_STOCKS_NUM, len(sheet[DATE_COLUMN]))):
        if sheet[DATE_COLUMN][idx].value is None:
            continue
        try:
            row_date = normalize_date(sheet[DATE_COLUMN][idx].value)
        except Exception as e:
            row_date = None

        if row_date == date:
            start_row = idx
        if start_row >= 0 and idx > start_row and row_date is not None:
            end_row = idx
            break
    if start_row >= 0 and end_row == -1:
        end_row = len(sheet[DATE_COLUMN])
    return start_row, end_row


def save_stock_info_to_excel(sheet, row, holding_stock_info, price_info):
    sheet["%s%d" % (HOLDING_VALUE_COLUMN, row)
          ] = holding_stock_info.current_value
    sheet["%s%d" % (OPENING_PRICE_COLUMN, row)
          ] = price_info.opening_price
    sheet["%s%d" % (CLOSING_PRICE_COLUMN, row)
          ] = price_info.closing_price
    sheet["%s%d" % (BUY_IN_PRICE_COLUMN, row)
          ] = holding_stock_info.buy_in_price
    sheet["%s%d" % (PROFIT_COLUMN, row)] = holding_stock_info.current_value - \
        holding_stock_info.buy_in_value
    sheet["%s%d" % (PROFIT_RATE_COLUMN, row)] = "%.2f" % (100.0*(holding_stock_info.current_value -
                                                                 holding_stock_info.buy_in_value) / holding_stock_info.buy_in_value) + "%"


def save_investment_info_to_excel(sheet, start_row, end_row, investment_info):
    sheet["%s%d" % (TOTAL_VALUE_COLUMN, start_row + 1)
          ] = investment_info.total_value
    sheet["%s%d" % (TOTAL_PROFIT_RATE_COLUMN, start_row + 1)
          ] = "%.2f" % (100.0*investment_info.profit_rate) + "%"
    sheet.merge_cells("%s%d:%s%d" % (TOTAL_VALUE_COLUMN,
                      start_row + 1, TOTAL_VALUE_COLUMN, end_row))
    sheet.merge_cells("%s%d:%s%d" % (TOTAL_PROFIT_RATE_COLUMN,
                      start_row + 1, TOTAL_PROFIT_RATE_COLUMN, end_row))


def draw_profit_history(wb, profit_history):
    if PROFIT_HISTORY_SHEET in wb.sheetnames:
        wb.remove(wb[PROFIT_HISTORY_SHEET])
    wb.create_sheet(PROFIT_HISTORY_SHEET)
    ws = wb[PROFIT_HISTORY_SHEET]
    ws["A1"] = "日期"
    ws["B1"] = "总资金"
    ws["C1"] = "收益率"
    for idx in range(0, len(profit_history)):
        ws["A%d" %
            (idx + 2)] = datetime.strptime(profit_history[idx][0], "%Y/%m/%d").date()
        ws["B%d" % (idx + 2)] = profit_history[idx][1]
        ws["C%d" % (idx + 2)] = profit_history[idx][2]

    total_rows = len(ws["A"])
    asset_line = LineChart()
    asset_line.title = "资金曲线"
    asset_line.style = 12
    asset_line.y_axis.title = "资金"
    asset_line.y_axis.crossAx = 500
    asset_line.x_axis = DateAxis(crossAx=100)
    asset_line.x_axis.title = "日期"
    asset_line.x_axis.number_format = "mm/dd"
    asset_line.x_axis.majorTimeUnit = "days"

    data = Reference(ws, min_col=2, min_row=1, max_col=2,
                     max_row=total_rows)
    asset_line.add_data(data, titles_from_data=True)

    dates = Reference(ws, min_col=1, min_row=2,
                      max_row=total_rows)
    asset_line.set_categories(dates)

    ws.add_chart(asset_line, "E10")


def process_daily_stock(investment_info, sheet, day, start_row_hint):
    cur_date = "%s-%02d" % (sheet.title, day)
    try:
        cur_date = datetime.strptime(cur_date, "%Y-%m-%d").strftime("%Y/%m/%d")
    except ValueError:
        return -1

    start_row, end_row = get_row_range(sheet, cur_date, start_row_hint)
    # print("### date: %s, start_row: %d, end_row: %d" %
    #       (cur_date, start_row, end_row))
    if start_row == -1:
        return end_row  # return previous row
    cur_holding_stocks = []
    for row in range(start_row, end_row):
        stock_code = str(sheet[STOCK_CODE_COLUMN][row].value).strip("\"\".")
        cur_holding_stocks.append(stock_code)

    # 先卖出不在当天自选股名单中的股票
    prev_holding_stocks = investment_info.get_holding_stocks()

    for st_code in prev_holding_stocks:
        if st_code not in cur_holding_stocks:

            # 如果股票从自选股中消失，则按照当天的开盘价卖出
            investment_info.sell_out_stock(
                st_code, cur_date)

    for row in range(start_row, end_row):
        stock_code = str(sheet[STOCK_CODE_COLUMN][row].value).strip("\"\".")
        stock_name = sheet[STOCK_NAME_COLUMN][row].value.strip("\"\".")
        # 当天的自选股信息
        stock = StockInfo(stock_code, stock_name)
        price_info = StockMarketHelper.get_stock_price_info(
            stock_code, cur_date)
        if price_info is None:
            print("Get empty stock info of stock %s, %s, at date %s" %
                  (stock_code, stock_name, cur_date))
            continue

        # 股票出现在自选股名单中，则按照当天的开盘价买入
        if not investment_info.is_holding_stock(stock_code):
            investment_info.buy_in_stock(
                stock, cur_date, price_info.opening_price)

        holding_stock = investment_info.get_holding_stock(stock_code)
        holding_stock.calculate_profit(cur_date, price_info.closing_price)
        save_stock_info_to_excel(sheet, row + 1, holding_stock, price_info)

    # 计算截止到当天整体的收益率
    investment_info.calculate_profit(cur_date)
    print("========== cur date: %s, investment info : %s ==========" %
          (cur_date, investment_info))
    save_investment_info_to_excel(
        sheet, start_row, end_row, investment_info)
    return end_row


def process_stock_sheet(investment_info, sheet):
    # 每月一个sheet, 处理sheet
    start_row_hint = 0
    for day in range(1, 32):
        start_row_hint = max(start_row_hint, process_daily_stock(
            investment_info, sheet, day, start_row_hint))


@click.command()
@click.option("--excel", default="2021-stocks.xlsx", help="excel file path")
@click.option("--year", default=2021, help="year of stock info")
def process_stock_excel(excel, year):
    investment_info = InvestmentInfo(INIT_TOTAL_VALUE)
    try:
        wb = openpyxl.load_workbook(excel)
    except Exception as e:
        raise RuntimeError("Failed to open excel file %s" % excel)

    for month in range(1, 13):
        sheet_name = "%d-%02d" % (year, month)
        try:
            ws = wb[sheet_name]
        except Exception as e:
            # raise RuntimeError("Failed to get sheet, error: %s", e)
            print("Failed to get sheet, error: %s" % e)
            continue
        if len(ws[DATE_COLUMN]) <= 2:  # 该sheet 为空
            continue
        process_stock_sheet(investment_info, ws)

    draw_profit_history(wb, investment_info.profit_history)
    wb.save(excel)


def test():
    wb = openpyxl.load_workbook("final.xlsx")
    print("### sheet names: %s" % wb.sheetnames)
    if PROFIT_HISTORY_SHEET in wb.sheetnames:
        wb.remove(wb[PROFIT_HISTORY_SHEET])
    wb.create_sheet(PROFIT_HISTORY_SHEET)
    ws = wb.get_sheet_by_name(PROFIT_HISTORY_SHEET)
    asset_line = LineChart()
    asset_line.title = "资金曲线"
    asset_line.style = 12
    asset_line.y_axis.title = "资金"
    asset_line.y_axis.crossAx = 500
    asset_line.x_axis = DateAxis(crossAx=100)
    asset_line.x_axis.title = "日期"
    asset_line.x_axis.number_format = 'd-mmm'
    asset_line.x_axis.majorTimeUnit = "days"

    total_row = len(ws["A"])
    data = Reference(ws, min_col=2, min_row=1,
                     max_row=total_row)
    asset_line.add_data(data, titles_from_data=True)

    dates = Reference(ws, min_col=1, min_row=2, max_row=total_row)
    asset_line.set_categories(dates)
    ws.add_chart(asset_line)
    wb.save("final1.xlsx")


if __name__ == "__main__":
    process_stock_excel()
    # test()
