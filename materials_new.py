#!/usr/bin/env python
# coding: utf-8
import pandas as pd
import constants as cn


# ### Read datafiles
def read_file(filename):
    return pd.read_csv(cn.DATAFILE_PATH + filename)


inv = read_file(cn.INVENTORY_EXPORT)
bl = read_file(cn.BACKLOG_EXPORT)
hfr = read_file(cn.HFR_EXPORT)
validate = read_file(cn.VALIDATION_DB)
translate = read_file(cn.TRANSLATION_DB)


# ### Set datafile dtypes
def set_float_columns(df):
    col = df.columns[1:]
    df[col] = df[col].astype(float)
    return df


inv = set_float_columns(inv)
bl = set_float_columns(bl)
hfr = set_float_columns(hfr)
translate = set_float_columns(translate)

# ### Set datafile column names
inv.columns = cn.INVENTORY_LABELS
bl.columns = cn.BACKLOG_LABELS
hfr.columns = cn.HFR_LABELS
validate.columns = cn.VALIDATE_LABELS
translate.columns = cn.TRANSLATE_LABELS


# ### Clean bl & hfr
def clean_sales_table(df):
    name = (list(df.columns))[1]
    df = df.loc[df[cn.PN] != " "]
    df = df.loc[df[cn.PN] != "LOT"]
    df = df.loc[df[cn.PN] != "MARK"]
    df = df.loc[df[cn.PN] != "MISC"]
    for row in df.index:
        df.loc[row, name] = df.loc[row, name] * df.loc[row, 'Factor']
    df = df.drop('Factor', axis=1)
    df = df.groupby([cn.PN], as_index=False).sum()
    return df


bl = clean_sales_table(bl)
hfr = clean_sales_table(hfr)

# ### Download schedule
tables = pd.read_html(cn.SCHEDULE_URL)

# ### Clean schedule
schedule = tables[0].fillna(0)
labels = [cn.PN] + list(schedule.loc[cn.DATES_ROW][cn.DATES_COL_START:])
schedule = schedule.drop(cn.SCHEDULE_USELESS_COLS, axis=1)
schedule = schedule[cn.SCHEDULE_USELESS_ROWS:]
schedule.columns = labels
schedule = set_float_columns(schedule)
schedule = schedule.groupby([cn.PN], as_index=False).sum()

# ### Translate schedule
schedule = schedule.merge(translate, how='left').fillna(0)
rows = schedule[schedule[cn.FAC] != 0].index
for col in schedule.columns[1:]:
    schedule.loc[rows, col] = schedule.loc[rows, col] * schedule.loc[rows, cn.FAC]
schedule = schedule.drop(cn.FAC, axis=1)


# ### Validate schedule
def fix_pn(toki, tli):
    if tli == 0:
        return toki
    if toki != tli:
        return tli
    else:
        return toki


schedule = validate.merge(schedule, how='right').fillna(0)
schedule['Part Number'] = schedule[['Part Number', 'Valid PN']].apply(
    lambda schedule: fix_pn(schedule['Part Number'], schedule['Valid PN']), axis=1)
schedule = schedule.drop('Valid PN', axis=1)

# ### Build materials
materials = pd.DataFrame()

materials[['Part Number', 'On Hand']] = inv[['Part Number', 'On Hand']]
materials = materials.merge(bl, how='left').fillna(0)
materials['Released'] = 0.0
materials = materials.merge(hfr, how='left').fillna(0)
materials['Released'] = materials['Backlog'] - materials['HFR']
materials['On Order'] = inv['On Order']
materials['T-Avail'] = materials['On Hand'] + materials['On Order'] - materials['Backlog']
materials['R-Avail'] = materials['T-Avail'] + materials['HFR']
materials['Reorder'] = inv['Reorder']
materials = materials.merge(schedule, how='left').fillna(0)
cols = materials.columns[1:]
materials[cols] = materials[cols].astype(int)

# ### Format and save files

df = materials
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
for i in range(9, len(df.columns)):
    ro_ws.write(0, i, ro.columns.values[i], ro_hdr_schedule)
    df_ws.write(0, i, df.columns.values[i], df_hdr_schedule)

ro_writer.save()
df_writer.save()

ro = ro.set_index('Part Number')
df = df.set_index('Part Number')
