import camelot
tables = camelot.read_pdf('0002.pdf', pages='1-end', flavor='lattice')
for i in range (0,tables.n):
    tables[i].to_excel("table_%d.xlsx" % i, index=False)