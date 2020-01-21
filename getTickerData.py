from openpyxl import Workbook
import requests
import time

def main():

    limit = 16000 #under 1.5s per ticker point average

    for y in range(2018, 2020):
        year = str(y)
        for m in range(1, 13):
            if m < 10:
                month = "0{}".format(m)
            else:
                month = str(m)
            a = []
            b = []
            c = []
            for d in range(1, 32):
                if d < 10:
                    day = "0{}".format(d)
                else:
                    day = str(d)
                date = "{}-{}-{}".format(year, month, day) #format: yyyy-mm-dd, can't have single digit for month or day
                url = "https://api.polygon.io/v2/ticks/stocks/nbbo/DG/{}?limit={}&apiKey={redactedAPIKey}".format(date, limit) #for DG
                mfile = requests.get(url).json()
                wb = Workbook()
                ws = wb.active
                for i in range(0, limit):
                    try:
                        print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(((mfile['results'][i]['t']) / 1000000000))))
                    except KeyError:
                        continue
                    a.append(time.strftime('%Y-%m-%d %H:%M:%S',
                                               time.localtime(((mfile['results'][i]['t']) / 1000000000))))
                    b.append(mfile['results'][i]['p'])
                    c.append(mfile['results'][i]['P'])

                ws.append(a)
                ws.append(b)
                ws.append(c)
                saveas = "DG-{}-prices.xlsx".format(date)
                wb.save(saveas)

if __name__ == '__main__':
        main()