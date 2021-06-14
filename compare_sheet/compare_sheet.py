import time
import xlwings as xw
import math

TIME_LIMIT = 100


def get_sheets() -> list:
    return ["Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"]


def compare_sheet(base_dir, compare_dir, sheet_name, countChanged, text):

    sht1 = xw.Book(base_dir).sheets[sheet_name]
    sht2 = xw.Book(compare_dir).sheets[sheet_name]

    sht1_row, sht1_col = sht1.used_range.shape
    sht2_row, sht2_col = sht2.used_range.shape

    row_count, col_count = max(sht1_row, sht2_row), max(sht1_col, sht2_col)

    count = row_count * col_count
    cur_count = 0

    for i in range(1, row_count + 1):
        for j in range(1, col_count + 1):
            cur_count += 1
            if sht1.range((i, j)).value != sht2.range((i, j)).value:

                text.emit(str((i, j)))
            countChanged.emit(math.ceil(cur_count / count * 100))

    text.emit("比较完成")


if __name__ == '__main__':
    base_dir = "D:/projects/demo/compare_sheet/files/a.xlsx"
    compare_dir = "D:/projects/demo/compare_sheet/files/b.xlsx"

    compare_sheet(base_dir, compare_dir, "Sheet1", "xxx")
