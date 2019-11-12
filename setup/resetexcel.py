#!/usr/bin/env python

import math
from pathlib import Path
from openpyxl import *

try:
	wb = load_workbook('../lyr.xlsx')
except: 
	wb = Workbook()
wb.remove(wb.active)
wb.create_sheet()
ws = wb.active
ws['A1'] = '0'
ws['E2'] = '0'
ws['E3'] = '0'
wb.save("../lyr.xlsx")