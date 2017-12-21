#! /usr/bin/env python
# -*- coding: utf-8 -*-

import time
import openpyxl
from openpyxl.styles import NamedStyle, Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from operator import itemgetter
import datetime
import pprint

            
def export_to_excel(data_list, worksheet_title, filename):
    wb = openpyxl.Workbook()

    ws = wb.worksheets[0]
    ws.title = worksheet_title

    row = 1
    for value in data_list:
        for i, data_point in enumerate(value):
            ws.cell(row=row, column=i+1).value = data_point

        row += 1

    wb.save(filename)
