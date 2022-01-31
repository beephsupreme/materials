materials = pd.read_excel('/Users/michael/projects/materials/data/materials.xlsx')
bom = pd.read_csv('/Users/michael/Desktop/tangers/bom.txt')
bom = bom.drop('UM_Multiplier', axis=1)
bom = bom.groupby('Part Number', as_index=False).sum()
materials = materials.merge(bom, how='right').fillna(0)
dates = list(materials.columns[9:])
cols_to_drop = ['T-Avail', 'R-Avail', 'Reorder']
cols_to_drop = cols_to_drop + dates
materials = materials.drop(cols_to_drop, axis=1)
materials['Available'] = materials['On Hand'] - materials['Qty Ordered'] - materials['Backlog']