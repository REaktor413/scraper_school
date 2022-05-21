import requests
from bs4 import BeautifulSoup
import lxml
import xlsxwriter


def decodeEmail(e):
    de = ""
    k = int(e[:2], 16)

    for i in range(2, len(e)-1, 2):
        de += chr(int(e[i:i+2], 16)^k)

    return de

headers = {'Mozilla/5.0':'(Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.82 Safari/537.36'}

url = f"https://www.bsupa.org.uk/sup_school/"

req = requests.get(url=url, headers=headers)
src = req.text

soup = BeautifulSoup(src, 'lxml')
result = []
for ultag in soup.find_all('ul', class_='list-sup-school'):
    for litag in ultag.find_all('li'):
        res = {}
        res['name'] = litag.find('h3').contents[0]
        res['adress'] = litag.find('address').contents[0]
        if litag.find('p').contents:
            res['phone'] = litag.find('p').contents[0] 
        else: res['phone'] = None
        try:
            res['email'] = decodeEmail(litag.find('a', class_='btn btn-blueongrey pushgap')['href'].split("#")[1])
        except:
            res['email'] = None
        if litag.find('a', class_='btn btn-blueongrey'):
            res['site'] = litag.find('a', class_='btn btn-blueongrey')['href']
        else: res['site'] = None
        print(res)
        result.append(res)


workbook = xlsxwriter.Workbook('school.xlsx')
worksheet = workbook.add_worksheet()
count = 0
for i in result:
    worksheet.write(f'A{count}', i["name"]) if i['name'] else worksheet.write(f'A{count}', None)
    worksheet.write(f'B{count}', i["adress"]) if i['adress'] else worksheet.write(f'A{count}', None)
    worksheet.write(f'C{count}', i["phone"]) if i['phone'] else worksheet.write(f'A{count}', None)
    worksheet.write(f'D{count}', i["email"]) if i['email'] else worksheet.write(f'A{count}', None)
    worksheet.write(f'E{count}', i["site"]) if i['site'] else worksheet.write(f'A{count}', None)
    count += 1

workbook.close()

print(f"Finish")
