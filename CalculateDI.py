#!/usr/bin/python3
# coding=utf-8
import os
from collections import defaultdict
import pandas as pd
from pandas import read_excel
import sys
from PyQt5.QtWidgets import QApplication, QMessageBox


class CalculateDI:
    """
    计算用例表中的用例DI值
    """
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.dict = self.read_excel()

    def read_excel(self):
        # 读取Excel文件的第一个工作表
        df = pd.read_excel(self.excel_path)
        # 将DataFrame转换为字典，键为列名，值为对应的列数据列表
        data_dict = df.to_dict(orient='list')
        # print(data_dict)

        return data_dict

    @property
    def calculate_di(self):
        try:
            serious = self.dict.get("严重程度", [])
            status = self.dict.get("Bug状态", [])

            if not serious or not status:
                raise KeyError("Excel文件缺少必要的列：'严重程度' 或 'Bug状态'")

            combined_dict = defaultdict(list)

            for key, value in zip(status, serious):
                combined_dict[key].append(value)

            combined_dict = dict(combined_dict)
            DI = 0
            for key, value in combined_dict.items():
                if "已关闭".__eq__(key):
                    continue
                DI = DI + 10 * value.count(1) + 3 * value.count(2) + value.count(3) + 0.5 * value.count(4)
            return DI
        except KeyError as ke:
            print(f"Key Error: {ke}")
            show_popup(f"错误：Excel 文件缺少必要的列：'{ke}'")
            sys.exit(1)


def is_excel_file(file_path):
    """
    检查给定的文件路径是否指向一个有效的Excel文件。

    :param file_path: 文件路径字符串
    :return: 如果是Excel文件则返回True，否则返回False
    """
    excel_extensions = ['.xls', '.xlsx', '.xlsm']
    _, ext = os.path.splitext(file_path.lower())
    return ext in excel_extensions


def show_popup(content):
    """
    QT弹框
    :param content:
    :return:
    """
    app = QApplication([])
    # 使用三个引号（即 ''' 或 """）定义的是多行字符串。这种语法允许你在字符串中包含换行符和空格等格式化信息，而不需要使用转义字符 \n 或 \t 来表示新行或制表符。
    app.setStyleSheet("""
        QMessageBox QLabel {
            min-width: 200px;
        }
        QMessageBox QPushButton {
            width: 100px;
            height: 30px;
        }
    """)
    msg_box = QMessageBox()
    msg_box.setText(content)
    msg_box.exec_()


if __name__ == '__main__':
    # excel_path = "/home/admin/Desktop/UOS AI/统信桌面操作系统 V25-Bug.xlsx"
    # cal = CalculateDI(excel_path)
    # # print(cal.calculate_di)
    # di_value = cal.calculate_di
    # show_popup(di_value)

    if len(sys.argv) < 2:
        print("Usage: python3 CalculateDI.py <path_to_excel>")
        sys.exit(1)

    excel_path = sys.argv[1]
    if not os.path.isfile(excel_path):
        show_popup("选择的文件不存在")
        sys.exit(1)

    if not is_excel_file(excel_path):
        show_popup("请选择一个有效的Excel文件")
        sys.exit(1)

    try:
        cal = CalculateDI(excel_path)
        di_value = cal.calculate_di
        show_popup(f'DI值是：{di_value}')
    except Exception as e:
        print(f"Unexpected error: {e}")
        show_popup(f"发生了一个意外错误：{str(e)}")