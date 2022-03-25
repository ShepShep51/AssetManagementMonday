import os
import sys
import json
import logging
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

column_lims = [4,27]

with open('STR Board Data.json', 'r') as json_file:
    bd_fmt = json.load(json_file)
with open('PropertyAbbreviations.JSON', 'r') as j_file:
    abrev = json.load(j_file)
with open('Performance Board Data.JSON', 'r') as jfile:
    financial_board_format = json.load(jfile)
with open('NCF_Board_Data.json', 'r') as infile:
    ncf_board_data = json.load(infile)

class Group:
    def __init__(self, board_id, name):
        self.board_id = board_id
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
class Pulse:
    def __init__(self, name, board_id, group_id,fmat,metrics):
        self.name = name
        self.board_id = board_id
        self.group_id = group_id
        self.fmat = fmat
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
        mutation = 'mutation {create_item (board_id: %s, item_name: "%s", group_id: "%s", column_values: "{%s}") {id}}' % (self.board_id, self.name, self.group_id, self.upload_string)
        mutation = {'query':mutation}
        r = requests.post(url=apiUrl, json=mutation, headers=headers)
        r = r.json()
        self.set_id(r['data']['create_item']['id'])
    ...

z = Group('2391151700', 'Testing Group Classes 1')
