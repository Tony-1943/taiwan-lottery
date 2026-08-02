import pandas as pd
import json
from copy import copy
from openpyxl import load_workbook
from TaiwanLottery import TaiwanLotteryCrawler
from docx import Document
from docx.shared import Pt
from datetime import datetime
# =====================================
# 日期格式
# =====================================

def roc_date(date_str):
    d = pd.to_datetime(date_str)
    return f"{d.year - 1911}.{d.month:02d}.{d.day:02d}"


def short_date(date_str):

    week_map = {
        0: "一",
        1: "二",
        2: "三",
        3: "四",
        4: "五",
        5: "六",
        6: "日"
    }

    d = pd.to_datetime(date_str)

    return f"{d.month:02d}/{d.day:02d}({week_map[d.weekday()]})"


# =====================================
# 獎號格式
# =====================================

# 539、大樂透、威力彩使用
def format_lottery_numbers(nums):

    return "  ".join(
        f"{int(n):02d}"
        for n in nums
    )


# 3星彩、4星彩使用
def format_star_numbers(nums):

    return "  ".join(
        str(int(n))
        for n in nums
    )

# =====================================
# 複製列格式
# =====================================

def copy_row_style(ws, source_row, target_row):

    for col in range(1, ws.max_column + 1):

        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)

        if src.has_style:

            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = copy(src.number_format)
            dst.protection = copy(src.protection)

    ws.row_dimensions[target_row].height = \
        ws.row_dimensions[source_row].height


# =====================================
# 清除資料
# =====================================

def clear_rows(ws, start_row=2):

    for r in range(start_row, 200):

        for c in range(1, ws.max_column + 1):

            ws.cell(r, c).value = None


# =====================================
# 抓取資料
# =====================================

def get_data():

    crawler = TaiwanLotteryCrawler()

    data = {
    "539": crawler.daily_cash()[:10],
    "威力彩": crawler.super_lotto()[:10],
    "大樂透": crawler.lotto649()[:10],
    "3星彩": crawler.lotto3d()[:10],
    "4星彩": crawler.lotto4d()[:10]
    }

    print("539:", len(data["539"]))
    print("威力彩:", len(data["威力彩"]))
    print("大樂透:", len(data["大樂透"]))
    print("3星彩:", len(data["3星彩"]))
    print("4星彩:", len(data["4星彩"]))

    return data


# =====================================
# 今彩539
# =====================================

def write_539(ws, data):

    clear_rows(ws)

    for r in range(3, 12):
        copy_row_style(ws, 2, r)

    for row, item in enumerate(reversed(data), start=2):

        ws.cell(row, 1).value = item["期別"]
        ws.cell(row, 2).value = roc_date(item["開獎日期"])
        ws.cell(row, 3).value = format_lottery_numbers(item["獎號"])


# =====================================
# 威力彩
# =====================================

def write_power(ws, data):

    clear_rows(ws)

    for r in range(3, 12):
        copy_row_style(ws, 2, r)

    for row, item in enumerate(reversed(data), start=2):

        ws.cell(row, 1).value = item["期別"]

        ws.cell(row, 2).value = roc_date(
            item["開獎日期"]
        )

        ws.cell(row, 3).value = format_lottery_numbers(
            item["第一區"]
        )

        ws.cell(row, 4).value = \
            f"{int(item['第二區']):02d}"


# =====================================
# 大樂透
# =====================================

def write_lotto649(ws, data):

    clear_rows(ws)

    for r in range(3, 12):
        copy_row_style(ws, 2, r)

    for row, item in enumerate(reversed(data), start=2):

        ws.cell(row, 1).value = item["期別"]

        ws.cell(row, 2).value = roc_date(
            item["開獎日期"]
        )

        ws.cell(row, 3).value = format_lottery_numbers(
            item["獎號"]
        )

        ws.cell(row, 4).value = \
            f"{int(item['特別號']):02d}"


# =====================================
# 3星彩
# =====================================

def write_3star(ws, data):

    clear_rows(ws)

    for r in range(3, 12):
        copy_row_style(ws, 2, r)

    for row, item in enumerate(reversed(data), start=2):

        ws.cell(row, 1).value = short_date(
            item["開獎日期"]
        )

        ws.cell(row, 2).value = format_star_numbers(
            item["獎號"]
        )


# =====================================
# 4星彩
# =====================================

def write_4star(ws, data):

    clear_rows(ws)

    for r in range(3, 12):
        copy_row_style(ws, 2, r)

    for row, item in enumerate(reversed(data), start=2):

        ws.cell(row, 1).value = short_date(
            item["開獎日期"]
        )

        ws.cell(row, 2).value = format_star_numbers(
            item["獎號"]
        )


# =====================================
# Excel
# =====================================

def create_excel(data):

    wb = load_workbook("Excel範本.xlsx")

    write_539(
        wb["539"],
        data["539"]
    )

    write_power(
        wb["威力彩"],
        data["威力彩"]
    )

    write_lotto649(
        wb["大樂透"],
        data["大樂透"]
    )

    write_3star(
        wb["3星"],
        data["3星彩"]
    )

    write_4star(
        wb["4星"],
        data["4星彩"]
    )

    wb.save("最新開獎紀錄.xlsx")

    print("Excel完成")

# =====================================
# Word 表格清空
# =====================================

def clear_table(table):

    for r in range(1, len(table.rows)):

        for cell in table.rows[r].cells:

            cell.text = ""


# =====================================
# Word 字型縮小
# =====================================

def resize_table_font(table):

    for row in table.rows:

        for cell in row.cells:

            for paragraph in cell.paragraphs:

                for run in paragraph.runs:

                    run.font.size = Pt(8)


# =====================================
# Word產生
# =====================================

def create_word(data):

    doc = Document("word 範本.docx")

    tables = doc.tables

    # 539
    table = tables[0]

    clear_table(table)

    for i, item in enumerate(reversed(data["539"][:10])):

        row = i + 1

        table.cell(row, 0).text = str(item["期別"])
        table.cell(row, 1).text = roc_date(item["開獎日期"])
        table.cell(row, 2).text = format_lottery_numbers(item["獎號"])

    resize_table_font(table)

    # 3星彩
    table = tables[1]

    clear_table(table)

    for i, item in enumerate(reversed(data["3星彩"][:10])):

        row = i + 1

        table.cell(row, 0).text = short_date(
            item["開獎日期"]
        )

        table.cell(row, 1).text = format_star_numbers(
            item["獎號"]
        )

    resize_table_font(table)

    # 威力彩
    table = tables[2]

    clear_table(table)

    for i, item in enumerate(reversed(data["威力彩"][:10])):

        row = i + 1

        table.cell(row, 0).text = str(item["期別"])

        table.cell(row, 1).text = roc_date(
            item["開獎日期"]
        )

        table.cell(row, 2).text = format_lottery_numbers(
            item["第一區"]
        )

        table.cell(row, 3).text = \
            f"{int(item['第二區']):02d}"

    resize_table_font(table)

    # 4星彩
    table = tables[3]

    clear_table(table)

    for i, item in enumerate(reversed(data["4星彩"][:10])):

        row = i + 1

        table.cell(row, 0).text = short_date(
            item["開獎日期"]
        )

        table.cell(row, 1).text = format_star_numbers(
            item["獎號"]
        )

    resize_table_font(table)

    # 大樂透
    table = tables[4]

    clear_table(table)

    for i, item in enumerate(reversed(data["大樂透"][:10])):

        row = i + 1

        table.cell(row, 0).text = str(item["期別"])

        table.cell(row, 1).text = roc_date(
            item["開獎日期"]
        )

        table.cell(row, 2).text = format_lottery_numbers(
            item["獎號"]
        )

        table.cell(row, 3).text = \
            f"{int(item['特別號']):02d}"

    resize_table_font(table)

    doc.save("最新開獎紀錄.docx")

    print("Word完成")




def create_json(data):

    result = {

        "update_time": str(pd.Timestamp.now()),

        "539": {
            "period": str(
                data["539"][0]["期別"]
            ),

            "date": roc_date(
                data["539"][0]["開獎日期"]
            ),

            "numbers": format_lottery_numbers(
                data["539"][0]["獎號"]
            )
        },

        "power": {

            "period": str(
                data["威力彩"][0]["期別"]
            ),

            "date": roc_date(
                data["威力彩"][0]["開獎日期"]
            ),

            "numbers": format_lottery_numbers(
                data["威力彩"][0]["第一區"]
            ),

            "special": f"{int(data['威力彩'][0]['第二區']):02d}"
        },

        "lotto649": {

            "period": str(
                data["大樂透"][0]["期別"]
            ),

            "date": roc_date(
                data["大樂透"][0]["開獎日期"]
            ),

            "numbers": format_lottery_numbers(
                data["大樂透"][0]["獎號"]
            ),

            "special": f"{int(data['大樂透'][0]['特別號']):02d}"
        },

        "3star": {

            "period": str(
                data["3星彩"][0]["期別"]
            ),

            "date": short_date(
                data["3星彩"][0]["開獎日期"]
            ),

            "numbers": format_star_numbers(
                data["3星彩"][0]["獎號"]
            )
        },

        "4star": {

            "period": str(
                data["4星彩"][0]["期別"]
            ),

            "date": short_date(
                data["4星彩"][0]["開獎日期"]
            ),

            "numbers": format_star_numbers(
                data["4星彩"][0]["獎號"]
            )
        }
    }

    with open(
        "data.json",
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            result,
            f,
            ensure_ascii=False,
            indent=4
        )

    print("JSON完成")
# =====================================
# 主程式
# =====================================

def main():

    print("開始抓取資料...")
######測試
    data = get_data()
    for game, records in data.items():
        print(f"{game}: {len(records)}")
######
    if all(len(v) == 0 for v in data.values()):

        print("本月尚無資料")

        print("保留上次產生的 Excel、Word、JSON")

        return
    
    print("539:", len(data["539"]))
    print("威力彩:", len(data["威力彩"]))
    print("大樂透:", len(data["大樂透"]))
    print("3星彩:", len(data["3星彩"]))
    print("4星彩:", len(data["4星彩"]))
    create_excel(data)
    create_word(data)
    print("全部完成")

    create_json(data)
if __name__ == "__main__":
    main()