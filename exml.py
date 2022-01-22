import pandas as pd

file = 'decmo.xlsx'

x1 = pd.ExcelFile(file)
print(x1.sheet_names)
print(x1.parse('january'))

x




