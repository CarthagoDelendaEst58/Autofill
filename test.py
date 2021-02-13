import pandas as pd
import json
import os.path

from FillFuncs import getAbcamData
from FillFuncs import getPubchemData

magento = pd.read_excel('database_sheets/magento_sept.xlsx')
sku = "0180237425"
# sku = "091760422"
sku = '0296033550'

data = getAbcamData(sku, magento)
# data = getPubchemData(sku, magento)

print(data)

# with open("Abcam/92-71-7.json", "r") as read_file:
# with open("Abcam/pls.json", "r") as read_file:
#     data = json.load(read_file)

# print(data['search_name'])

# print(os.path.exists("Abcam/Phosphate Buffered Saline (PBS), Dulbecco's formula, powder, w/o Ca, Mg, .json"))
# path = "Abcam/PhosphateBufferedSaline(PBS)Dulbecco'sformulapowderwoCaMg.json"
# path = "Abcam/Phosphate Buffered Saline (PBS), Dulbecco's formula, powder, wo Ca, Mg, .json"
# open(path, 'w')