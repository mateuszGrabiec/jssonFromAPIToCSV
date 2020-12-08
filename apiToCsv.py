import argparse
import requests
import json
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

ap = argparse.ArgumentParser()
ap.add_argument("-u", "--url", required=True,
	help="input url with request")
ap.add_argument("-d", "--data", required=True,
	help="input which json property from response contain array")
ap.add_argument("-n", "--name", required=False,
	help="name of output file")
args = vars(ap.parse_args())

response = requests.get(args['url'])
response = json.loads(response.text)
collection = response[args['data']]

def getKeys(obj):
    keys=[]
    for key in obj.keys():
        if isinstance(obj[key], dict)== False:
           keys.append(key)
    return keys

jsonKeys = getKeys(collection[0])

ws.append(jsonKeys)

for col in collection:
    row = []
    for k in jsonKeys:
        r = col[k]
        if r:
            row.append(r)
        else:
            row.append('')
    ws.append(row)

saveName = 'out.xlsx'
if args['name']:
    saveName = args['name']+'.xlsx'

wb.save(saveName)
