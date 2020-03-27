import camelot, os, fnmatch, time
import pandas as pd
from progress.bar import Bar
import xlsxwriter

directory = './dataset'
# print(len(fnmatch.filter(os.listdir(directory), '*.pdf'))) = 4371

start_time = time.time()
with Bar('Processing', max=4284) as bar: #max=4371
    for filename in sorted(os.listdir(directory)):
        pdfpath=os.path.join(directory, filename) #./testset/ + 0001.pdf
        tables = camelot.read_pdf(pdfpath, pages='1-end', flavor='lattice')
        frames = pd.DataFrame()
        for i in range (0,tables.n):
            frames = frames.append(tables[i].df, ignore_index=True)
        savepath = os.path.join('./xlsx', os.path.splitext(filename)[0] + '.xlsx') #./test/ + 0001.xlsx
        frames.to_excel(savepath, engine='xlsxwriter', header=False, index=False)
        #frames.to_csv(savepath, header=False, index=False, encoding='utf_8')
        bar.next()
print("--- %s seconds ---" % (time.time() - start_time))