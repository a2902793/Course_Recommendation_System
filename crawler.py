import urllib
import urllib.request
import os
from progress.bar import Bar

with Bar('Processing', max=20) as bar:
    for x in range(1, 9999) :
        id = str(x).zfill(4)
        url = "http://ap09.emis.tku.edu.tw/108_2/108_2_%s.PDF" % id
        try: # http 200
            resp = urllib.request.urlopen(url)
        except urllib.error.HTTPError as e: # http 404
            pass
        else:
            fullfilename = os.path.join('./pdf/', id + '.pdf')
            urllib.request.urlretrieve(url, fullfilename)
        bar.next()