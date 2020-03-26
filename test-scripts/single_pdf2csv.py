import camelot, os, fnmatch, time
import pandas as pd

tables = camelot.read_pdf('./testset/0001.pdf', pages='1-end', flavor='lattice')
frames = pd.DataFrame()
frames = frames.append(tables[0].df, ignore_index=True)
frames = frames.append(tables[1].df, ignore_index=True)
frames = frames.append(tables[2].df, ignore_index=True)
frames.to_csv("c.csv", header=False, index=False, encoding='utf_8')
#frames.to_excel("c.xlsx", header=False, index=False)