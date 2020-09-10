import xlwt
import pandas as pd
from shanghai_stock_exchange.stock_data_deal import deal_day_data, get_all_stock_day_data, \
    deal_day_data_4_auto_invest_plan_a, deal_day_data_4_auto_invest_plan_b, deal_day_data_4_auto_invest_plan_c
from shanghai_stock_exchange.utils.date_utils import date_add_for_year
import sys
sys.setrecursionlimit(5000000)

# 定投计划分析数据excel导出
# stock_code_list [{'code': '000001', 'name':'上证指数', 'module': 'SZ'}]
def export_auto_invest_plan_analyze_data(stock_code_list):

    for stock_code in stock_code_list:
        total_auto_invest_plan_list = []
        for plan_index in range(3):
            cur_total_auto_invest_plan_dict = {}
            if plan_index == 0:
                cur_total_auto_invest_plan_dict = deal_day_data_4_auto_invest_plan_a(stock_code['code'], stock_code['module'])
            if plan_index == 1:
                cur_total_auto_invest_plan_dict = deal_day_data_4_auto_invest_plan_b(stock_code['code'], stock_code['module'])
            if plan_index == 2:
                cur_total_auto_invest_plan_dict = deal_day_data_4_auto_invest_plan_c(stock_code['code'], stock_code['module'])

            cur_total_auto_invest_plan_list = cur_total_auto_invest_plan_dict['auto_invest_ok_list']
            if cur_total_auto_invest_plan_list is not None:
                total_auto_invest_plan_list.extend(cur_total_auto_invest_plan_list)
        # 整理数据格式准备导出
        export_excel_4_auto_invest_plan_data(total_auto_invest_plan_list,stock_code['name'])


# 股票历史涨跌数据excel导出
def export_stock_data(stock_code_list=None):
    # beg_time = date_add_for_year(None, -5, '%Y-%m-%d')
    week_day_list = []
    if stock_code_list is None:
        stock_data_list = get_all_stock_day_data(None)
        if stock_data_list is None:
            return
        for week_day_dict in stock_data_list:
            week_day_list_sub = dispose_one_stock_week_day_obj(week_day_dict)
            if week_day_list_sub is None:
                continue
            week_day_list.extend(week_day_list_sub)
        return
    for stock_code in stock_code_list:
        week_day_dict = deal_day_data(stock_code['code'], stock_code['module'])
        week_day_list_sub = dispose_one_stock_week_day_obj(week_day_dict)
        if week_day_list_sub is None:
            continue
        week_day_list.extend(week_day_list_sub)
    export_excel(week_day_list)


# 通过网络请求获取的单个week_day_dict组装week_day_list
def dispose_one_stock_week_day_obj(week_day_dict):
    week_day_list = []
    for i in range(7):
        week_day_obj = week_day_dict['week_day_' + str(i + 1)]
        if week_day_obj is None:
            continue
        week_day_obj['stock_code'] = week_day_dict['stock_code']
        week_day_obj['stock_name'] = week_day_dict['stock_name']
        week_day_list.append(week_day_obj)

    return week_day_list


def export_excel(export):
    # 将字典列表转换为DataFrame
    pf = pd.DataFrame(list(export))
    # 指定字段顺序
    order = ['stock_code', 'stock_name', 'week_day', 'up_times', 'total_up_range', 'down_times', 'total_down_range']
    pf = pf[order]
    # 将列名替换为中文
    columns_map = {
        'stock_name': '股票名称',
        'stock_code': '股票编号',
        'week_day': '周几',
        'up_times': '上涨次数',
        'total_up_range': '累计上涨幅度',
        'down_times': '下跌次数',
        'total_down_range':'累计下跌幅度'
    }
    pf.rename(columns=columns_map, inplace=True)
    # 指定生成的Excel表格名称
    file_path = pd.ExcelWriter('stock_data_export_01_.xlsx')
    # 替换空单元格
    pf.fillna(' ', inplace=True)
    # 输出
    pf.to_excel(file_path, encoding='utf-8', index=False)
    # 保存表格
    file_path.save()


def export_excel_4_auto_invest_plan_data(export,file_name_suffix):
    if export is None:
        return
    # 将字典列表转换为DataFrame
    pf = pd.DataFrame(list(export))
    # 指定字段顺序
    order = ['stock_name', 'stock_code', 'stock_plan', 'start_deal_date', 'end_deal_date', 'take_days', 'auto_invest_ok_count', 'auto_invest_count','auto_invest_ok_rate']
    pf = pf[order]
    # 将列名替换为中文
    columns_map = {
        'stock_name': '股票名称',
        'stock_code': '股票编号',
        'stock_plan': '定投计划',
        'start_deal_date': '开始交易时间',
        'end_deal_date': '结束交易时间',
        'take_days':'持续天数',
        'auto_invest_ok_count': '定投成功次数',
        'auto_invest_count':'定投总数',
        'auto_invest_ok_rate': '定投成功比率'
    }
    pf.rename(columns=columns_map, inplace=True)
    # 指定生成的Excel表格名称
    file_path = pd.ExcelWriter('stock_auto_invest_data_export' + file_name_suffix + '.xlsx')
    # 替换空单元格
    pf.fillna(' ', inplace=True)
    # 输出
    pf.to_excel(file_path, encoding='utf-8', index=False)
    # 保存表格
    file_path.save()

    # cur_ok_auto_invest['stock_plan'] = stock_plan
    # cur_ok_auto_invest['start_deal_date'] = date_str
    # cur_ok_auto_invest['end_deal_date'] = cur_auto_invest_plan_result[1]
    # cur_ok_auto_invest['auto_invest_ok_count'] = auto_invest_ok_count
    # last_line_analyze_data['auto_invest_count'] = auto_invest_count
    # last_line_analyze_data['auto_invest_ok_rate']


# 将分析完成的列表导出为excel表格
# export_stock_data(
#     [{'code': '000001', 'module': 'SZ'}, {'code': '399001', 'module': 'SC'}, {'code': '399006', 'module': 'CY'}])


# 将分析完成的定投数据列表导出为excel表格
export_auto_invest_plan_analyze_data(
    [{'code': '000001', 'name':'上证指数', 'module': 'SZ'}])

# , {'code': '399001', 'name':'深证成指', 'module': 'SC'}
