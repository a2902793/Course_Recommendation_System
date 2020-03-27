import camelot, os, fnmatch, time
import pandas as pd
import xlsxwriter

tables = camelot.read_pdf('./dataset/0087.pdf', pages='1-end', flavor='lattice')
frames = pd.DataFrame()
frames = frames.append(tables[0].df, ignore_index=True)
frames = frames.append(tables[1].df, ignore_index=True)
frames = frames.append(tables[2].df, ignore_index=True)
#frames.to_csv("c.csv", header=False, index=False, encoding='utf_8')
frames.to_excel("0087.xlsx", engine='xlsxwriter', header=False, index=False)