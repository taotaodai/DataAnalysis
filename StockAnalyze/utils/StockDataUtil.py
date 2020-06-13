
#特殊处理
def isST(stock_name):
    return stock_name.find('ST') > 0
#退市股
def isDelist(stock_name):
    return stock_name.find('退') > 0

