# -*- coding: utf-8 -*-
import openpyxl
import csv
import os

# ==================== 可配置参数 ====================
SRC_FILE = '课表.xlsx'      # 输入文件
DST_FILE = '转换结果.csv'     # 输出文件
SHEET_INDEX = 0               # 工作表索引，0 表示第一个表
# ===================================================

def convert_xlsx_to_csv(src_file=SRC_FILE, dst_file=DST_FILE, sheet_index=SHEET_INDEX):
    """把 src_file 指定的 xlsx 转成 dst_file 指定的 csv。

    既支持单独运行本脚本（走默认的 课表.xlsx -> 转换结果.csv），
    也支持被 test.py 导入后传入它自己输出的文件路径。
    返回 True 表示转换成功，False 表示源文件不存在。
    """
    if not os.path.exists(src_file):
        print(f"[错误] 找不到文件：{src_file}")
        return False

    print(f"正在读取 {src_file} ...")
    wb = openpyxl.load_workbook(src_file, data_only=True)
    ws = wb.worksheets[sheet_index]

    # 1. 收集所有合并单元格的信息，做成一个字典 { (行, 列): 值 }
    merged_data = {}
    for rng in ws.merged_cells.ranges:
        # 获取合并区域左上角的值
        top_left_value = ws.cell(rng.min_row, rng.min_col).value
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                merged_data[(row, col)] = top_left_value

    # 2. 遍历所有单元格，构建二维数组
    data = []
    for row in ws.iter_rows():
        row_data = []
        for cell in row:
            # 优先取合并单元格字典里的值，如果没有则取当前单元格的值
            val = merged_data.get((cell.row, cell.column), cell.value)
            # 处理 None 值，并去除首尾空白
            if val is None:
                row_data.append('')
            else:
                row_data.append(str(val).strip())
        data.append(row_data)

    # 3. 写入 CSV 文件
    # encoding='utf-8-sig' 是防止用 Excel 打开 CSV 时中文乱码的关键
    try:
        with open(dst_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(data)
    except PermissionError:
        print(f"[错误] 无法写入 {dst_file}，该文件可能正被 Excel / WPS 打开。")
        print("请关闭该文件后重新运行。")
        return False

    print(f"转换完成！合并单元格已自动填充。")
    print(f"输出文件：{dst_file}")
    return True

if __name__ == '__main__':
    convert_xlsx_to_csv()