import os
import sys
import json
import logging
import time

import requests
import calendar
import pandas as pd
import tkinter as tk
import xlwings as xw
import datetime as dt
import openpyxl as pxl
import tkinter.filedialog

apiKey = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjEyNDU2NTYzMywidWlkIjoyNDYzODQ0MSwiaWFkIjoiMjAyMS0wOS0xNFQxMzozMzozNS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NDMzODY1LCJyZ24iOiJ1c2UxIn0.pWJkEFC3rxdwtYOMgbyvmDoI6xjRYxqZVbIdBmUfa6k"
apiUrl = "https://api.monday.com/v2"
headers = {"Authorization": apiKey}
conversions = {"LOF REIT - Fund 2":"lof2",
               "LF3 REIT - Fund 3":"lf3",
               "ACCEL II":"Accel 2",
               "Legendary Lodging VAB QOZ":"VABQOZ"}
column_lims = [4,27]
test_path = r'C:\Users\dshepard\PycharmProjects\AssManUpload\venv\LF3 PFolio Dashboard - DEC - FINAL.xlsx'


with open('STR Board Data.json', 'r') as json_file:
    bd_fmt = json.load(json_file)
with open('PropertyAbbreviations.JSON', 'r') as j_file:
    abrev = json.load(j_file)
with open('Performance Board Data.JSON', 'r') as jfile:
    financial_board_format = json.load(jfile)
with open('NCF_Board_Data.json', 'r') as infile:
    ncf_board_data = json.load(infile)


class Board:
    def __init__(self, j_file, fund, idx):
        self.j_file = j_file
        self.fund = fund
        self.fund_boards = self.j_file[self.fund]
        self.idx = idx
        self.board_key = self.fund_boards[self.idx].keys()
        self.boardID()
        # self.board_id = '2391151700'
    def multiBoards(self):
        if type(self.idx) == 'list':
            for i in self.idx:
                ...
        ...
    def boardID(self):
        for k in self.board_key:
            self.board_id = self.fund_boards[self.idx][k]['id']
            self.board_format = self.fund_boards[self.idx][k]['column_data']


class Group(Board):
    def __init__(self, board, name):
        self.board_id = board.board_id
        self.board_format = board.board_format
        self.name = name
        self.create_mutation()
        self.create_group()
    def create_mutation(self):
        mutation = 'mutation {create_group (board_id: %s, group_name: "%s"){id}}' %(self.board_id,self.name)
        self.mutation = {'query':mutation}
    def create_group(self):
        r = requests.post(url=apiUrl, json=self.mutation, headers=headers)
        r = r.json()
        self.group_id = r['data']['create_group']['id']


    ...
class Pulse(Group):
    def __init__(self, group, name, metrics):
        self.board_id = group.board_id
        self.group_id = group.group_id
        self.fmat = group.board_format
        self.name = name
        self.metrics = metrics
        self.data_string()
        self.create_pulse()
        
    def set_id(self,item_id):
        self.id = item_id
        
    def data_string(self):
        format = r'\"\": \"\", '
        end_format = r'\"\": \"\"'
        final = []
        for i in range(len(self.fmat)):
            if i != len(self.fmat) - 1:
                final.append(format[:2] + self.fmat[i]['id'] + format[2:8] + str(self.metrics[i]) + format[8:])

            else:
                final.append(end_format[:2] + self.fmat[i]['id'] + end_format[2:8] + str(self.metrics[i]) + end_format[8:])
        final_string = ''
        for i in final:
            final_string = final_string + i
        self.upload_string = final_string

    
    def create_pulse(self):
        mutation = 'mutation {create_item (board_id: %s, item_name: "%s", group_id: "%s", column_values: "{%s}") {id}}' % (self.board_id, self.name, self.group_id, self.upload_string) # , column_values: "{%s}"    , self.upload_string
        mutation = {'query':mutation}
        r = requests.post(url=apiUrl, json=mutation, headers=headers)
        r = r.json()

    ...

def data_pull(worksheet_obj, column_limits):
    fund_dict = {'LOF REIT - Fund 2': [], 'LF3 REIT - Fund 3': [], 'Legendary Lodging VAB QOZ': [], 'ACCEL II': []}
    last_row = worksheet_obj.range(worksheet_obj.cells.last_cell.row, 2).end('up').row
    for row in range(1, last_row):
        if worksheet_obj.range(row, 2).value in fund_dict.keys():
            fund_dict[worksheet_obj.range(row, 2).value].append(row +1)
            for i in range(row+1,row + 21):
                if worksheet_obj.range(i, 3).value is None or worksheet_obj.range(i, 3).value == 'Total LF3 Core Portfolio':
                    fund_dict[worksheet_obj.range(row, 2).value].append(i - 1)
                    break
    data = {}
    for k in fund_dict.keys():
        r_list = fund_dict[k]
        data[k] = {}
        for r in range(r_list[0],r_list[1]+1):
            name = worksheet_obj.range(r,column_limits[0]-1).value
            data[k][name] = []
            for c in range(column_limits[0],column_limits[1]):
                if worksheet_obj.range(r,c).value is not None and worksheet_obj.range(8,c).value != "% Chg Rank":
                    if type(worksheet_obj.range(r,c).value) == float:
                        data[k][name].append(float("{0:.2f}".format(worksheet_obj.range(r,c).value)))
                    else:
                        data[k][name].append(worksheet_obj.range(r,c).value)


    return data

def Main():
    op = options()
    file_path = browser()

    if op['Upload Type'] == '1':
        wb = xw.Book(file_path)
        tab_to_open = tabSelect(timeframe=op['Timeframe'], workbook_object=wb)
        ws = wb.sheets[str(tab_to_open)]
        fund_dict = data_pull(worksheet_obj=ws)
        reit_two_lims = fund_dict['LOF REIT - Fund 2']
        reit_three_lims = fund_dict['LF3 REIT - Fund 3']
        vab_lims = fund_dict['Legendary Lodging VAB QOZ']
        accel_lims = fund_dict['ACCEL II']
        if op['Fund'] == '2':
            gt_row = reit_two_lims[1]
            fund_abrev = abrev['lof2']
            if op['Timeframe'] == '2':
                gt_board = Board(bd_fmt,'lof2',0)
                gt_board.board_format = gt_board.board_format[:-1]
                gt_group = Group(gt_board,tab_to_open)
                gt_pulse = Pulse(gt_group,)

with open('test_data.json', 'r') as infile:
    testing = json.load(infile)

# GOOD TO GO
# ts = time.time()
# wb = xw.Book(test_path)
# ws = wb.sheets['Jan 16 - Jan 22']
# x = data_pull(ws,column_lims)
# ts1 = time.time()
#
# for fund in x.keys():
#     if 'VAB' in fund:
#         prop_board = Board(j_file=bd_fmt,fund=conversions[fund],idx=0)
#         prop_group = Group(prop_board,'Jan 16 - Jan 22')
#     else:
#         prop_board = Board(j_file=bd_fmt,fund=conversions[fund],idx=3)
#         gt_board = Board(j_file=bd_fmt,fund=conversions[fund],idx=2)
#         prop_group = Group(prop_board,'Jan 16 - Jan 22')
#         gt_group = Group(gt_board,'Jan 16 - Jan 22')
#         for entry in x[fund].keys():
#             if "Total" in entry:
#                 Pulse(gt_group,entry,x[fund][entry])
#             else:
#                 Pulse(prop_group,entry,x[fund][entry])
# with open('test_data.json', 'w') as outfile:
#     json.dump(x,outfile,indent=2)
# END SECTION

with open('MasterBoardData.json','r') as infile:
    master = json.load(infile)
k = []
for i in range(len(financial_board_format['Accel II'])):

    for j in financial_board_format['Accel II'][i]:
        print(j)
        master['ACCEL II']['Performance'][j] = financial_board_format['Accel II'][i][j]



with open('MasterBoardData.json','w') as outfile:
    json.dump(master,outfile,indent=2)