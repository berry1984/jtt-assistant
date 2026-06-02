#!/usr/bin/env python3
"""
将航乐 .xls 模板转换为 .xlsx 格式（一次性转换脚本）。
openpyxl 不支持读取旧版 .xls，所以用 xlrd 读取后再用 openpyxl 创建。

用法：
  python3 convert_template_xls_to_xlsx.py
"""
import os
import sys

THIS_DIR = os.path.dirname(os.path.abspath(__file__))

# ── 要转换的模板 ──
TEMPLATES = [
    ('航乐-客户名称 客户单号 英国发票模板9.9更新.xls', '航乐-客户名称 客户单号 英国发票模板9.9更新.xlsx'),
    ('航乐-客户单号- 欧州发票模板2.26更新.xls', '航乐-客户单号- 欧州发票模板2.26更新.xlsx'),
]

def convert_xls_to_xlsx(xls_path, xlsx_path):
    """用 xlrd 读 .xls，用 openpyxl 写 .xlsx"""
    import xlrd
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter

    xls_wb = xlrd.open_workbook(xls_path, formatting_info=True)
    xlsx_wb = Workbook()

    # 删掉默认 Sheet
    xlsx_wb.remove(xlsx_wb.active)

    for si in range(xls_wb.nsheets):
        xls_ws = xls_wb.sheet_by_index(si)
        ws = xlsx_wb.create_sheet(title=xls_ws.name)

        # ── 列宽 ──
        for c in range(xls_ws.ncols):
            ci = xls_ws.colinfo_map.get(c)
            if ci:
                # xlrd 列宽单位 ≈ 1/256 字符宽度，openpyxl 用字符宽度
                ws.column_dimensions[get_column_letter(c+1)].width = ci.width / 256 * 1.2

        # ── 行高 + 行数据 ──
        for r in range(xls_ws.nrows):
            ri = xls_ws.rowinfo_map.get(r)
            if ri:
                ws.row_dimensions[r+1].height = ri.height / 20  # twips → points

            for c in range(xls_ws.ncols):
                cell_type = xls_ws.cell_type(r, c)
                cell_value = xls_ws.cell_value(r, c)

                if cell_type == 0:  # Empty
                    continue

                xl_cell = xlsx_wb.active if False else None  # placeholder

                if cell_type == 1:  # Text
                    ws.cell(row=r+1, column=c+1, value=str(cell_value))
                elif cell_type == 2:  # Number
                    ws.cell(row=r+1, column=c+1, value=cell_value)
                elif cell_type == 3:  # Date
                    ws.cell(row=r+1, column=c+1, value=cell_value)
                elif cell_type == 4:  # Boolean
                    ws.cell(row=r+1, column=c+1, value=bool(cell_value))
                elif cell_type == 5:  # Error
                    pass  # skip errors
                elif cell_type == 6:  # Blank (has formatting but no value)
                    pass
                else:
                    ws.cell(row=r+1, column=c+1, value=str(cell_value))

        # ── 合并单元格 ──
        for merge in xls_ws.merged_cells:
            # xlrd: (rlo, rhi, clo, chi) inclusive-exclusive
            rlo, rhi, clo, chi = merge
            # openpyxl: 1-indexed, inclusive-inclusive
            ws.merge_cells(
                start_row=rlo+1, start_column=clo+1,
                end_row=rhi, end_column=chi
            )

    xlsx_wb.save(xlsx_path)
    print(f'✅ 已转换: {xls_path} → {xlsx_path}')

if __name__ == '__main__':
    for xls_name, xlsx_name in TEMPLATES:
        xls_path = os.path.join(THIS_DIR, xls_name)
        xlsx_path = os.path.join(THIS_DIR, xlsx_name)
        if not os.path.exists(xls_path):
            print(f'⚠️  源文件不存在: {xls_path}')
            continue
        convert_xls_to_xlsx(xls_path, xlsx_path)
