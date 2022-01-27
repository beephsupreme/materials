# -*- coding: utf-8 -*-
import inventory
import shipping
import sales
import report
import pandas as pd
import constants as cn
from datetime import date


def run():
    data = inventory.build()
    hfr = sales.build(cn.HFR_EXPORT)
    backlog = sales.build(cn.BACKLOG_EXPORT)
    schedule = shipping.build(data[cn.PN].tolist())
    df = report.build(data, schedule, backlog, hfr)

    df_writer = pd.ExcelWriter("./data/materials.xlsx", engine="xlsxwriter")
    df.to_excel(df_writer, index=False, sheet_name="Materials")
    df_wb = df_writer.book
    df_ws = df_writer.sheets["Materials"]

    ro = df[(df['Reorder'] > df['T-Avail'])]
    ro_writer = pd.ExcelWriter("./data/reorder.xlsx", engine="xlsxwriter")
    ro.to_excel(ro_writer, index=False, sheet_name="Report")
    ro_wb = ro_writer.book
    ro_ws = ro_writer.sheets["Report"]

    ro_hdr_pn = ro_wb.add_format({
        'bold': True,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'left',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    ro_hdr_body = ro_wb.add_format({
        'bold': True,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'right',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    ro_hdr_schedule = ro_wb.add_format({
        'bold': True,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'center',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    ro_pn = ro_wb.add_format({
        'bold': False,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'left',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    ro_body = ro_wb.add_format({
        'bold': False,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'right',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    ro_schedule = ro_wb.add_format({
        'bold': False,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'center',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    df_hdr_pn = df_wb.add_format({
        'bold': True,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'left',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    df_hdr_body = df_wb.add_format({
        'bold': True,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'right',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    df_hdr_schedule = df_wb.add_format({
        'bold': True,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'center',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    df_pn = df_wb.add_format({
        'bold': False,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'left',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    df_body = df_wb.add_format({
        'bold': False,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'right',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    df_schedule = df_wb.add_format({
        'bold': False,
        'text_wrap': False,
        'valign': 'bottom',
        'align': 'center',
        'fg_color': '#FFFFFF',
        'font_color': '#000000',
        'border': 1,
        'border_color': 'E0E0E0'
    })

    ro_ws.set_column("A:A", 20, ro_pn)
    ro_ws.set_column("B:I", 10, ro_body)
    ro_ws.set_column("J:Z", 10, ro_schedule)
    df_ws.set_column("A:A", 20, df_pn)
    df_ws.set_column("B:I", 10, df_body)
    df_ws.set_column("J:Z", 10, df_schedule)

    ro_ws.write(0, 0, ro.columns.values[0], ro_hdr_pn)
    df_ws.write(0, 0, df.columns.values[0], df_hdr_pn)

    for i in range(1, 9):
        ro_ws.write(0, i, ro.columns.values[i], ro_hdr_body)
        df_ws.write(0, i, df.columns.values[i], df_hdr_body)
    for i in range(9, 9 + cn.SHIPPING_WIDTH):
        ro_ws.write(0, i, ro.columns.values[i], ro_hdr_schedule)
        df_ws.write(0, i, df.columns.values[i], df_hdr_schedule)

    ro_writer.save()
    df_writer.save()

    ro = ro.set_index('Part Number')
    df = df.set_index('Part Number')

    cn.HEADER = ["Part Number", "On Hand", "Backlog", "Released",
                 "HFR", "On Order", "T-Avail", "R-Avail", "Reorder"]
    cn.SHIPPING_DATES = []
    cn.SHIPPING_WIDTH = 0

    return df, ro


if __name__ == "__main__":
    run()
