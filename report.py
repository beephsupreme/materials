# -*- coding: utf-8 -*-
import pandas as pd
import constants as cn


def build(data, schedule, bl, hfr):
    # setup header and report dataframe
    header = cn.HEADER
    header.extend(cn.SHIPPING_DATES)
    df = pd.DataFrame(columns=header)
    # populate materials with inventory data
    df[cn.PN] = data[cn.PN]
    df[cn.OH] = data[cn.OH]
    df[cn.OO] = data[cn.OO]
    df[cn.RO] = data[cn.RO]
    # zero-out remaining columns
    df[[cn.BL, cn.RLS, cn.HFR, cn.TA, cn.RA]] = 0
    df[cn.SHIPPING_DATES] = 0
    # loop through df and fill fields according to part number status
    for row in df.index:
        key = df.loc[row, cn.PN]
        if key in bl:
            df.loc[row, cn.BL] = bl[key]
            df.loc[row, cn.RLS] = bl[key]
        if key in hfr:
            df.loc[row, cn.HFR] = hfr[key]
            df.loc[row, cn.RLS] = bl[key] - hfr[key]
        if key in schedule:
            for col in range(cn.SHIPPING_WIDTH):
                df.loc[row, cn.SHIPPING_DATES[col]] = (schedule[key])[col]
            # for index, date in enumerate(cn.SHIPPING_DATES):
            #     df.loc[row, date] = (schedule[key])[index]
        # calculate t-avail & r-avail
        df.loc[row, cn.TA] = df.loc[row, cn.OH] + df.loc[row, cn.OO] - df.loc[row, cn.BL]
        df.loc[row, cn.RA] = df.loc[row, cn.TA] + df.loc[row, cn.HFR]

    # convert numeric columns to int
    int_cols = [cn.OH, cn.BL, cn.RLS, cn.HFR, cn.OO, cn.TA, cn.RA, cn.RO]
    int_cols.extend(cn.SHIPPING_DATES)
    for col in int_cols:
        df[col] = df[col].astype(int)

    return df
