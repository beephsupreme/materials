# -*- coding: utf-8 -*-
import pandas as pd
import constants as cn


def build():
    data = pd.read_csv(cn.DATAFILE_PATH + cn.INVENTORY_AV_EXPORT)
    data.columns = [cn.PN, cn.OH, cn.OO, cn.RO]
    return data
