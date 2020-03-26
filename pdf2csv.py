import camelot, os, fnmatch, time
import pandas as pd
from progress.bar import Bar

directory = './dataset'
# print(len(fnmatch.filter(os.listdir(directory), '*.pdf'))) = 4371

start_time = time.time()
with Bar('Processing', max=4371) as bar:
    for filename in sorted(os.listdir(directory)):
        pdfpath=os.path.join(directory, filename) #./testset/ + 0001.pdf
        tables = camelot.read_pdf(pdfpath, pages='1-end', flavor='lattice')
        frames = pd.DataFrame()
        for i in range (0,tables.n):
            frames = frames.append(tables[i].df, ignore_index=True)
        savepath = os.path.join('./csv', os.path.splitext(filename)[0] + '.xlsx') #change to ./test/ + 0001.pdf
        #frames.to_excel(savepath, header=False, index=False)
        frames.to_csv(savepath, header=False, index=False, encoding='utf_8')
        bar.next()
print("--- %s seconds ---" % (time.time() - start_time))