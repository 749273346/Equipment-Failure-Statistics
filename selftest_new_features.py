import os
import tempfile

import openpyxl

from auto_fill_defects import DefectProcessor


def main():
    with tempfile.TemporaryDirectory() as d:
        excel_path = os.path.join(d, "test.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=3, column=1, value="1")
        ws.cell(row=3, column=2, value="模板")
        wb.save(excel_path)

        long_path = os.path.join(d, "a" * 80, "b" * 80, "demo.docx")
        row_data = ["", "广州", "类型A", "地点A", "", "", "", "", "", "", "", "", "", long_path]
        row_data2 = ["", "深圳", "类型B", "地点B", "", "", "", "", "", "", "", "", "", long_path]

        p = DefectProcessor(log_callback=lambda *_: None)
        wrote = p._write_rows_to_excel(excel_path, [row_data])
        if wrote != 1:
            raise RuntimeError(f"写入条数异常: {wrote}")

        wrote2 = p._write_rows_to_excel(excel_path, [row_data2])
        if wrote2 != 1:
            raise RuntimeError(f"追加写入条数异常: {wrote2}")

        wb2 = openpyxl.load_workbook(excel_path)
        ws2 = wb2.active

        if ws2.max_row != 5:
            raise RuntimeError(f"行数异常: {ws2.max_row}")

        cell_path = ws2.cell(row=4, column=14).value
        if str(cell_path) != long_path:
            raise RuntimeError("源文件路径未写入第14列")

        if not ws2.column_dimensions["N"].hidden:
            raise RuntimeError("第14列未隐藏")

        h = ws2.row_dimensions[4].height
        if h is not None and h > 150:
            raise RuntimeError(f"行高异常: {h}")

        wrote_refresh = p._write_rows_to_excel(excel_path, [row_data, row_data2], overwrite=True)
        if wrote_refresh != 2:
            raise RuntimeError(f"刷新写入条数异常: {wrote_refresh}")

        wb3 = openpyxl.load_workbook(excel_path)
        ws3 = wb3.active
        if ws3.max_row != 5:
            raise RuntimeError(f"刷新后行数异常: {ws3.max_row}")
        if str(ws3.cell(row=4, column=1).value) != "1":
            raise RuntimeError("刷新后序号异常")

    print("OK")


if __name__ == "__main__":
    main()
