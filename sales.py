# -*- coding: utf-8 -*-
import pandas as pd
import constants as cn


def build(filename):
    df = pd.read_csv(cn.DATAFILE_PATH + filename)
    df.columns = [cn.PN, cn.BL, cn.MP]
    df.loc[:, cn.BL] = df.loc[:, cn.BL] * df.loc[:, cn.MP]
    df.drop(cn.MP, axis=1, inplace=True)
    df = df.groupby([cn.PN]).sum()
    dictionary = (df.to_dict())[cn.BL]
    return dictionary
