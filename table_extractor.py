import camelot
tables = camelot.read_pdf('0001.pdf', pages='1-end', flavor='lattice')
for i in range (0,tables.n):
    #filename = "table_%d.xlsx" % i
    tables[i].to_excel("table_%d.xlsx" % i, index=False)