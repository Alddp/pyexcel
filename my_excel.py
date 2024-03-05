from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string
import csv


class Excel_robot:
    """docstring for Excel_robot."""

    def __init__(self, filename: str, data_only: bool, read_only: bool = False):
        self.filename = filename
        self.wb = load_workbook(self.filename, data_only=data_only, read_only=read_only)
        self.max_row: int | None = None
        self.max_col: int | None = None

    def get_sheet_names(self):
        """
        打印所有的sheet名称
        """
        sheet_names = self.wb.sheetnames
        print("sheets: " + str(sheet_names))
        return sheet_names

    def get_data(
        self,
        sheet_name: str,
        max_col_string: str,
        max_row: int | None = None,
        contains_first_line=False,
    ):
        # 读取sheet所有内容写入data
        """
        sheet_name: 要读取的列表
        """
        self.max_row = max_row
        self.max_col = column_index_from_string(max_col_string)
        ws = self.wb[sheet_name]

        data = []
        # 遍历所有cell
        for row in ws.iter_rows(
            min_row=2 if not contains_first_line else 1,
            max_col=self.max_col,
            max_row=self.max_row,
        ):
            data.append([cell.value for cell in row])

        with open(f"{sheet_name}.csv", "w", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(value for value in data)

        return sheet_name

    def read_csv(self, filename: str):
        data = []
        with open(filename + ".csv", "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            # 处理data的None值, 如果为存在None则忽略当行
            for row in reader:
                # print(row)
                if None in row or "" in row or len(row) < 1:
                    continue
                filtered_row = [value for value in row]
                data.append(filtered_row)

        return data

    def sorting_column(
        self,
        data: list,
        col_string: str,
        desc=False,
        convert_to_int=False,
    ):
        """
        根据列排序
        """

        # sorted_col_index = column_index_from_string(col_string) - 1
        # sorted_data = sorted(data, key=lambda x: x[sorted_col_index], reverse=desc)
        # return sorted_data

        sorted_col_index = column_index_from_string(col_string) - 1
        if convert_to_int:
            data.sort(key=lambda x: int(x[sorted_col_index]), reverse=desc)
        else:
            data.sort(key=lambda x: x[sorted_col_index], reverse=desc)

    def sorting_data(self, data):
        self.sorting_column(data, "B")
        self.sorting_column(data, "A")
        self.sorting_column(data, "E", desc=True, convert_to_int=True)
        return data

    def write_to_wb(
        self,
        sheet_name: str,
        sorted_data,
        start_row: int,
    ):
        wb = self.wb
        ws = wb[sheet_name]
        for i, row in enumerate(
            ws.iter_rows(
                min_row=start_row,
                max_row=self.max_row + 1,
                max_col=self.max_col,
            )
        ):
            for j, cell in enumerate(row):
                try:
                    cell.value = sorted_data[i][j]
                except Exception as e:
                    # print(e)
                    pass


if __name__ == "__main__":

    # 创建对象
    robot = Excel_robot("./红白榜数据汇总.xlsx", data_only=True)
    # 获取所有sheet
    sheet_names = robot.get_sheet_names()

    # 导出sheet的内容到csv,并返回当前sheet名称
    sheet_male = robot.get_data(
        sheet_names[-1],
        max_col_string="E",
        max_row=426,
    )
    print(f"{sheet_male}.csv dump ok...")

    sheet_female = robot.get_data(
        sheet_names[-2],
        max_col_string="E",
        max_row=381,
    )
    print(f"{sheet_female}.csv dump ok...")

    # 得到处理过的data
    processed_male_data = robot.read_csv(sheet_male)
    processed_female_data = robot.read_csv(sheet_female)

    # 排序
    sorted_male_data = robot.sorting_data(processed_male_data)  # 男生排序
    sorted_female_data = robot.sorting_data(processed_female_data)  # 女生排序

    # 写入Excel
    robot.write_to_wb(sheet_male, sorted_data=sorted_male_data, start_row=2)
    robot.write_to_wb(sheet_female, sorted_data=sorted_female_data, start_row=2)
    robot.wb.save(robot.filename + "sorted.xlsx")
    print(f"{robot.filename} 已经保存")
