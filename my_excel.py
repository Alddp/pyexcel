from openpyxl import load_workbook, Workbook, cell
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string
import csv
import pyuca


class Excel_robot:
    """docstring for Excel_robot."""

    def __init__(
        self,
        filename: str,
        data_only: bool,
        filename_summery: str = str | None,
        read_only: bool = False,
    ):
        self.filename = filename
        self.filename_summery = filename_summery
        self.wb = load_workbook(self.filename, data_only=data_only, read_only=read_only)
        self.max_row: int | None = None
        self.max_col: int | None = None
        self.summery_wb = load_workbook("./红白榜结果汇总.xlsx")

        self.red_male_sheet = self.summery_wb["红榜（男）"]
        self.red_female_sheet = self.summery_wb["红榜（女）"]
        self.green_sheet = self.summery_wb["较差"]
        self.yellow_sheet = self.summery_wb["白榜"]
        self.red_male_start_row = 4
        self.red_female_start_row = 4
        self.green_start_row = 4
        self.yellow_start_row = 4

    def get_sheet_names(self, wb: Workbook):
        """
        打印所有的sheet名称
        """
        sheet_names = wb.sheetnames
        print("sheets: " + str(sheet_names), end="\n\n")
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

        # 创建一个Unicode排序器
        collator = pyuca.Collator()

        # sorted_col_index = column_index_from_string(col_string) - 1
        # sorted_data = sorted(data, key=lambda x: x[sorted_col_index], reverse=desc)
        # return sorted_data

        sorted_col_index = column_index_from_string(col_string) - 1
        if convert_to_int:
            data.sort(key=lambda x: int(x[sorted_col_index]), reverse=desc)
        else:
            # data.sort(key=lambda x: x[sorted_col_index], reverse=desc)
            data.sort(
                key=lambda x: collator.sort_key(x[sorted_col_index]), reverse=True
            )

    def sorting_data(self, data) -> list:
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
                max_row=self.max_row,
                max_col=self.max_col,
            )
        ):
            for j, cell in enumerate(row):
                try:
                    cell.value = sorted_data[i][j]
                    if j == column_index_from_string("E") - 1:
                        self.fill_color(cell=cell)
                except Exception as e:
                    # print(e)
                    pass

    # 将分数分为各个阶段
    def divide_score(self, score: int):
        if score >= 100:
            return "red"
        if 60 <= score < 65:
            return "green"
        if 0 < score < 60:
            return "yellow"
        return None

    def fill_color(self, cell: cell):
        red = PatternFill("solid", fgColor="00FF0000")
        green = PatternFill("solid", fgColor="92D050")
        yellow = PatternFill("solid", fgColor="00FFFF00")

        value = int(cell.value)
        # if value >= 100:
        #     cell.fill = red
        # if 60 <= value < 65:
        #     cell.fill = green
        # if value < 60:
        #     cell.fill = yellow
        stage = self.divide_score(value)
        if stage == "red":
            cell.fill = red

        elif stage == "green":
            cell.fill = green

        elif stage == "yellow":
            cell.fill = yellow
        else:
            pass

    def summery(
        self,
        sorted_data,
        sheet_name,
    ):
        # 读取csv,将符合条件的行复制到当前行
        self.red_male_start_row = 4
        self.red_female_start_row = 4
        self.green_start_row = 4
        self.yellow_start_row = 4

        for i, row in enumerate(sorted_data):
            score = int(row[-1])
            stage = self.divide_score(score)
            if stage is None:
                continue
            # 判断类别
            if stage == "red":
                if sheet_name == "男生排序":
                    ws = self.red_male_sheet
                    self.red_male_start_row += 1
                    current_row = self.red_male_start_row
                if sheet_name == "女生排序":
                    ws = self.red_female_sheet
                    self.red_female_start_row += 1
                    current_row = self.red_female_start_row
            elif stage == "green":
                ws = self.green_sheet
                self.green_start_row += 1
                current_row = self.green_start_row
            elif stage == "yellow":
                ws = self.yellow_sheet
                self.yellow_start_row += 1
                current_row = self.yellow_start_row

            for j, value in enumerate(row):
                # 将当行写入汇总表
                ws.cell(row=current_row, column=j + 1, value=value)


if __name__ == "__main__":

    # 创建对象
    robot = Excel_robot(
        filename="./红白榜数据汇总.xlsx",
        filename_summery="./红白榜结果汇总.xlsx",
        data_only=True,
    )
    # 获取所有sheet
    sheet_names = robot.get_sheet_names(robot.wb)

    # 导出sheet的内容到csv,并返回当前sheet名称
    sheet_male = robot.get_data(
        sheet_names[-1],
        max_col_string="E",
        max_row=None,
    )
    print(f"{sheet_male}.csv dump ok...\n")

    sheet_female = robot.get_data(
        sheet_names[-2],
        max_col_string="E",
        max_row=None,
    )
    print(f"{sheet_female}.csv dump ok...\n")

    # 得到处理过的data
    processed_male_data = robot.read_csv(sheet_male)
    processed_female_data = robot.read_csv(sheet_female)

    # 排序
    sorted_male_data = robot.sorting_data(processed_male_data)  # 男生排序
    sorted_female_data = robot.sorting_data(processed_female_data)  # 女生排序

    # 写入Excel并填充颜色
    robot.write_to_wb(sheet_male, sorted_data=sorted_male_data, start_row=2)
    robot.write_to_wb(sheet_female, sorted_data=sorted_female_data, start_row=2)
    robot.wb.save(robot.filename + "sorted.xlsx")
    print(f"{robot.filename} 已经保存\n")

    # TODO: 将各个阶段的分数写入汇总表
    # 加载汇总表的sheets
    # '白榜', '较差', '红榜（男）', '红榜（女）'
    # robot.summery_wb = load_workbook("./红白榜结果汇总.xlsx")

    # red_male_sheet = robot.summery_wb["红榜（男）"]
    # red_female_sheet = robot.summery_wb["红榜（女）"]
    # green_sheet = robot.summery_wb["较差"]
    # yellow_sheet = robot.summery_wb["白榜"]
    # red_male_start_row = 4
    # red_female_start_row = 4
    # green_start_row = 4
    # yellow_start_row = 4

    robot.summery(sorted_data=sorted_male_data, sheet_name=sheet_male)
    robot.summery(sorted_data=sorted_female_data, sheet_name=sheet_female)

    robot.summery_wb.save("./汇总.xlsx")
    print("ok")
