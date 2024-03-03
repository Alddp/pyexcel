from openpyxl import load_workbook

wb = load_workbook("./红白榜数据汇总.xlsx", data_only=True)
# options = {"": "", "": ""}
cho = "男生排序"

ws = wb[cho]


# 获取男生排序表的数据
def get_data(ws):
    data = []

    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=471)):
        data.append([])

        for cell in row:
            data[i].append(cell.value)
    return data


def sorting_score(data, desc=True):
    # 根据"E"分数排序
    sorted_data = sorted(data, key=lambda x: x[-1], reverse=desc)
    return sorted_data


def sorting_building_no(data, desc=False):
    # 根据"A"楼号排序
    sorted_building_no = sorted(data, key=lambda x: x[0], reverse=desc)
    return sorted_building_no


def sorting_dormitory_no(data, desc=False):
    # 根据"B"宿舍号排序
    sorted_dormitory_no = sorted(data, key=lambda x: x[1], reverse=desc)
    return sorted_dormitory_no


def write_sorted_data(sorted_data, ws):
    # 写入EXCEL
    for i, row in enumerate(sorted_data, start=1):
        for j, value in enumerate(row, start=1):
            ws.cell(row=i + 1, column=j, value=value)


def main_sorting(data):
    # 主要排序:分数 降序
    # 次要排序:楼号, 宿舍号 升序
    # 先进行次要排序

    sorted_building_data = sorting_building_no(data)
    sorted_dormitory_data = sorting_dormitory_no(sorted_building_data)
    sorted_data = sorting_score(sorted_dormitory_data)
    return sorted_data


if __name__ == "__main__":
    data = get_data(ws)
    sorted_data = main_sorting(data)
    write_sorted_data(sorted_data, ws)
    # TODO: 添加表格颜色

    wb.save("./sorted.xlsx")
    print("OK")
