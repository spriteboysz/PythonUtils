#! /usr/bin/env python
# coding=utf-8
"""
Author: Deean
Date: 2023-09-23 16:04
FileName: 
Description:LeetcodeStatistics.py 
"""
import os
from collections import defaultdict

import xlwings

language = ["java", "py", "rb", "cpp", "c", "go", "js", "cs", "rs", "kt", "sql"]
leetcode_path = r'D:\02_CODE'
excel_file = r'D:\08_PYTH\PyUtils\LeetCode记录.xlsx'


def walk_data():
    record1 = defaultdict(lambda: defaultdict(int))
    record2 = defaultdict(lambda: defaultdict(int))
    for root, dirs, files in os.walk(leetcode_path):
        for file in filter(lambda f: f.split(".")[-1] in language, files):
            suffix = file.split(".")[-1]
            name = file[:-(len(suffix) + 1)]
            # 基础算法题
            if file.startswith("P"):
                name = name.split(".")[0]
                record1[name][suffix] += 1
            # 面试题
            elif file.startswith('面'):
                name = file[:9]
                record2[name][suffix] += 1
            elif file.startswith('M0'):
                name = "面试题 " + file[3:5] + "." + file[7:9]
                record2[name][suffix] += 1
            # LCP LCR LCS
            elif file.startswith("LC"):
                if "." in name:
                    name = name.split(".")[0]
                elif file.startswith("LCP") or file.startswith("LCS"):
                    name = file[:3] + " " + file[5:7]
                elif file.startswith("LCR"):
                    name = "LCR " + file[4:7]
                record2[name][suffix] += 1
            # 剑指题
            # elif file[0] in ['O', '剑']:
            #     record2[name][suffix] += 1
    return record1, record2


def to_grid(record):
    rows = sorted(list(record.keys()))
    m, n = len(rows), len(language)
    grid = [["" for _ in range(n + 2)] for _ in range(m + 2)]

    grid[0][0] = "题目列表"
    grid[0][1] = "统计"
    for j, lan in enumerate(language):
        grid[0][j + 2] = lan
    for i, row in enumerate(rows):
        grid[i + 2][0] = row
        grid[i + 2][1] = str(len(record[row]))
    for i, row in enumerate(rows):
        for j, lan in enumerate(language):
            if record[row][lan] > 0:
                grid[i + 2][j + 2] = "√"
    return grid


def to_excel(grids):
    with xlwings.App(visible=True, add_book=False) as app:
        workbook = app.books.add()
        for grid in grids:
            sheet = workbook.sheets.add()
            sheet.range("A1").value = grid
            sheet.autofit()
            sheet.range("C3").select()
            workbook.app.api.ActiveWindow.FreezePanes = True
            # sheet.range("A1").api.AutoFilter(1, "<>", True)
        workbook.save(excel_file)
        workbook.close()


if __name__ == '__main__':
    record1, record2 = walk_data()
    grid1, grid2 = to_grid(record1), to_grid(record2)
    to_excel([grid1, grid2])
