import glob
import json
import os
import re
import sys

import pandas as pd

import xlrd


def get_row_col(cell_str, zn_col_offset=0):
    col_r = r"[0-9]{1,}"
    row_r = r"[A-Z]{1,}"
    return (
        int(re.findall(col_r, cell_str)[0]) - 1 + zn_col_offset,
        ord(re.findall(row_r, cell_str)[0]) - ord("A"),
    )


if __name__ == "__main__":
    # 运行py文件时指定
    path = sys.path[0]
    # 编译为excel时指定
    # path = sys.argv[0]
    print(path)
    if os.path.isfile(path):
        path = os.path.dirname(path)
    files = glob.glob(path + os.sep + "source" + os.sep + "*.xls")
    # 按修改时间排序
    files = sorted(files, key=lambda x: os.stat(x).st_mtime)
    print(files)
    # 存储数据
    data_dict = {}
    for index, file in enumerate(files):
        # 读取excel数据
        print(file)
        data = xlrd.open_workbook(file)
        table = data.sheets()[0]
        # # 身份证号码
        # print(table.cell_value(get_row_col("C6")[0], get_row_col("C6")[1]))
        # # 本月子女支出(1000,看百分比)
        # print(table.cell_value(get_row_col("G11")[0], get_row_col("G11")[1]))
        # print(table.cell_value(get_row_col("G14")[0], get_row_col("G14")[1]))
        # # 继续教育支出(400每条)
        # print(table.cell_value(get_row_col("G17")[0], get_row_col("G17")[1]))
        # # 住房贷款(1000)
        # print(table.cell_value(get_row_col("G23")[0], get_row_col("G23")[1]))
        # # 住房租金支出(1500)
        # print(table.cell_value(get_row_col("G32")[0], get_row_col("G32")[1]))
        # # 赡养老人支出(2000封顶,按百分比)
        # print(table.cell_value(get_row_col("G38")[0], get_row_col("G38")[1]))
        zn_col_offset = 0        
        if "子女" in table.cell_value(get_row_col("A11")[0], get_row_col("A11")[1]):
            zn_col_offset = 0
        if "子女" in table.cell_value(get_row_col("A15")[0], get_row_col("A15")[1]):
            zn_col_offset = 4
        if "子女" in table.cell_value(get_row_col("A19")[0], get_row_col("A19")[1]):
            zn_col_offset = 8
        
        # 被赡养人定位
        bsyr_col_offset = 0
        if "被赡养人" in table.cell_value(get_row_col("A41", zn_col_offset)[0], get_row_col("A41", zn_col_offset)[1]):
            bsyr_col_offset = 2
           
        data_dict[index] = {
            "身份证号码": table.cell_value(get_row_col("C6")[0], get_row_col("C6")[1]),
            "本月子女支出1": table.cell_value(
                get_row_col("G14", 0)[0], get_row_col("G14", 0)[1]
            ),
            "本月子女支出2": table.cell_value(
                get_row_col("G14", 4)[0], get_row_col("G14", 4)[1]
            ),
            "本月子女支出3": table.cell_value(
                get_row_col("G14", 8)[0], get_row_col("G14", 8)[1]
            ),
            # 根据子女支出的条数, 确定偏移量
            "继续教育支出1": table.cell_value(
                get_row_col("G17", zn_col_offset)[0], get_row_col("G17", zn_col_offset)[1]
            ),
            "继续教育支出2": table.cell_value(
                get_row_col("G18", zn_col_offset)[0], get_row_col("G18", zn_col_offset)[1]
            ),
            "住房贷款": table.cell_value(
                get_row_col("G24", zn_col_offset)[0], get_row_col("G24", zn_col_offset)[1]
            ),
            "住房租金支出": table.cell_value(
                get_row_col("C34", zn_col_offset)[0], get_row_col("C34", zn_col_offset)[1]
            ),
            "赡养老人支出": table.cell_value(
                get_row_col("G38", zn_col_offset)[0], get_row_col("G38", zn_col_offset)[1]
            ),
            "婴幼儿一": table.cell_value(
                get_row_col("E49", zn_col_offset+bsyr_col_offset)[0], get_row_col("E49", zn_col_offset+bsyr_col_offset)[1]
            ),
            "婴幼儿二": table.cell_value(
                get_row_col("E49", zn_col_offset+bsyr_col_offset+2)[0], get_row_col("E49", zn_col_offset+bsyr_col_offset+2)[1]
            ),
        }
    print(data_dict)
    data_pd = pd.read_json(json.dumps(data_dict), orient="index")
    # 排序
    data_pd = data_pd[
        [
            "身份证号码",
            "本月子女支出1",
            "本月子女支出2",
            "本月子女支出3",
            "继续教育支出1",
            "继续教育支出2",
            "住房贷款",
            "住房租金支出",
            "赡养老人支出",
            "婴幼儿一",
            "婴幼儿二",
        ]
    ]
    data_pd["身份证号码"] = data_pd["身份证号码"].astype("str")
    data_pd.to_excel("{filename}.xlsx".format(filename="text"), index=False)
