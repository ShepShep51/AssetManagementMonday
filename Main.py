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

def browser():
    root = tk.Tk()
    #root.withdraw()

    currdir = os.getcwd()
    tempdir = tkinter.filedialog.askopenfilename(parent=root, initialdir=currdir, title='Select File You Want To Convert')
    root.destroy()
    root.mainloop()
    return tempdir


def dataUploadOption():
    print("What data would you like to upload?")
    upload = input(
        "Press [1] to upload STR Data, [2] to upload Financial Data, [3] to upload NCF Data, or [0] to return to home ")
    if upload == '1':
        print("Which fund are you uploading data for?")
        fund = input(
            "Press [1] for all funds, [2] for LOF2, [3] for LF3, [4] for Accel 2, [5] for VAB QOZ, or [6] to return to home ")
        print("What is the timeframe for the data you're uploading?")
        tf = input("Press [1] for Weekly, or press [2] for Monthly ")
        choices = {'Upload Type': upload,
                   'Fund': fund,
                   'Timeframe': tf}
        logging.info('Upload Type: {a}, Fund: {b}, Timeframe: {c}'.format(a=upload, b=fund, c=tf))
        return choices
        ...
    elif upload == '2':
        print("Which fund are you uploading?")
        fund = input("Press [2] for LOF2, [3] for LF3, [4] for Accel 2, or [5] for VAB QOZ: ")
        choices = {'Upload Type': upload,
                   'Fund': fund,
                   'Timeframe': None}
        logging.info('Upload Type: {a}, Fund: {b}, Timeframe: {c}'.format(a=upload, b=fund, c=None))
        return choices
        ...
    elif upload == '3':
        print("Which fund are you uploading?")
        fund = input("Press [2] for LOF2, or [3] for LF3 ")
        choices = {'Upload Type': upload,
                   'Fund': fund,
                   'Timeframe': None}
        logging.info('Upload Type: {a}, Fund: {b}, Timeframe: {c}'.format(a=upload, b=fund, c=None))
        return choices
        ...
    elif upload == '0':
        options()
        ...
    else:
        print('That is not a valid command, please try again')
        dataUploadOption()
    return

def abrevDataOption():
    with open('PropertyAbbreviations.JSON', 'r') as test:
        test_list = json.load(test)
    print("Which fund are you updating?")
    fund = input("Press [1] for LF3, or Press [2] for Accel 2: ")
    if fund == '1':
        name = input("Copy and paste the name of the property as it appears in the Excel sheet from Finance ")
        for i in test_list['lf3']:
            for k in i:
                if name == k:
                    print("It looks like that property is already in the database")
                    kekw = input('Press [1] to try again, press [2] to return to the homepage')
                    if kekw == '1':
                        abrevDataOption()
                    else:
                        options()
        abrev = input("Now type what you would like the abbreviation to be, then press enter ")
        addition = {name: abrev}
        test_list["lf3"].append(addition)
        logging.info('Added property: {a}, property abbreviation: {b}'.format(a=name, b=abrev))
        with open('PropertyAbbreviations.JSON', 'w') as json_file:
            json.dump(test_list, json_file, indent=2)
    elif fund == '2':
        name = input("Copy and paste the name of the property as it appears in the Excel sheet from Finance ")
        for i in test_list['lf3']:
            for k in i:
                if name == k:
                    print("It looks like that property is already in the database")
                    kekw = input('Press [1] to try again, press [2] to return to the homepage')
                    if kekw == '1':
                        abrevDataOption()
                    else:
                        options()
        abrev = input("Now type what you would like the abbreviation to be, then press enter ")
        addition = {name: abrev}
        logging.info('Added property: {a}, property abbreviation: {b}'.format(a=name, b=abrev))
        test_list["Accel 2"].append(addition)
        with open('PropertyAbbreviations.JSON', 'w') as json_file:
            json.dump(test_list, json_file, indent=2)

def options():
    print("Would you like to upload data or access property abbreviations?")
    nuts = input("Press [1] to upload data, press [2] to access propery abbreviations, or [0] to exit ")
    if nuts == "1": # Uploading data to Monday.com
        choices = dataUploadOption()
        return choices
    elif nuts == "2": # Updating or adding property abbreviations
        abrevDataOption()
        options()
    elif nuts == '0':
        logging.info('Exiting System')
        sys.exit()
    else:
        print('That is not a valid command')
        options()
        return


def data_pull(worksheet_obj):
    fund_dict = {'LOF REIT - Fund 2': [], 'LF3 REIT - Fund 3': [], 'Legendary Lodging VAB QOZ': [], 'ACCEL II': []}
    last_row = worksheet_obj.range(worksheet_obj.cells.last_cell.row, 2).end('up').row
    for row in range(1, last_row):
        if worksheet_obj.range(row, 2).value in fund_dict.keys():
            fund_dict[worksheet_obj.range(row, 2).value].append(row +1)
            for i in range(row+1,row + 21):
                if worksheet_obj.range(i, 3).value is None or worksheet_obj.range(i, 3).value == 'Total LF3 Core Portfolio':
                    fund_dict[worksheet_obj.range(row, 2).value].append(i - 1)
                    break
    return fund_dict


def newPulse(board_ident,group_ident,name_of_pulse,data_string):
    try:
        mutation = 'mutation {create_item (board_id: %s, item_name: "%s", group_id: "%s", column_values: "{%s}") {id}}' % (board_ident, name_of_pulse, group_ident,data_string)
        data = {'query': mutation}
        r = requests.post(url=apiUrl,json=data,headers=headers)
        response_data = r.json()
        item_id = response_data['data']['create_item']['id']
        logging.info('Pulse Created: {a}, Item ID: {b}'.format(a=name_of_pulse, b=item_id))
        return item_id
    except Exception as e:
        logging.exception(e)


def groupCreate(board_ident, group_name):
    try:
        mutation = 'mutation {create_group (board_id: %s, group_name: "%s"){id}}' % (board_ident, group_name)
        data = {'query': mutation}
        r = requests.post(url=apiUrl, json=data, headers=headers)
        response = r.json()
        print(response)
        group_id = response['data']['create_group']['id']
        logging.info('Group {a} created on board {b}. ID:{c}'.format(a=group_name,b=board_ident,c=group_id))
        return group_id
    except Exception as e:
        logging.exception(e)


def tabSelect(timeframe, workbook_object):
    if timeframe =='2':
        tabs = []
        indexes = []
        for i in range(0, len(workbook_object.sheets)):
            if workbook_object.sheets[i].name.find('Monthly') > 0:
                tabs.append(workbook_object.sheets[i].name)
                indexes.append(i)
        tabs = tabs[-5:]
        indexes = indexes[-5:]
        print("Select what data you'd like to upload:")
        tab_name = input("Press [1] for %s, [2] for %s, [3] for %s, [4] for %s, or [5] for %s " % (
        tabs[0], tabs[1], tabs[2], tabs[3], tabs[4]))
        tab_name = int(tab_name) - 1
        tab_index = indexes[tab_name]
        tab_name = tabs[tab_name]
        logging.info('Tab Selected: {a}'.format(a=tab_name))
        return tab_name
        ...
    else:
        tabs = []
        indexes = []
        for i in range(0, len(workbook_object.sheets)):
            tabs.append(workbook_object.sheets[i].name)
            indexes.append(i)
        tabs = tabs[-5:]
        indexes = indexes[-5:]
        print("Select what data you'd like to upload:")
        tab_name = input("Press [1] for %s, [2] for %s, [3] for %s, [4] for %s, or [5] for %s " % (
            tabs[0], tabs[1], tabs[2], tabs[3], tabs[4]))
        tab_name = int(tab_name) - 1
        tab_index = indexes[tab_name]
        tab_name = tabs[tab_name]
        logging.info('Tab Selected: {a}'.format(a=tab_name))
        return tab_name
        ...


def stringBuilder(formatting, metrics):
    format = r'\"\": \"\", '
    end_format = r'\"\": \"\"'
    final = []
    for i in range(len(formatting)):
        if i != len(formatting)-1:
            final.append(format[:2]+formatting[i]['id']+format[2:8]+str(metrics[i])+format[8:])
        else:
            final.append(end_format[:2] + formatting[i]['id'] + end_format[2:8] + str(metrics[i]) + end_format[8:])
    final_string = ''
    for i in final:
        final_string = final_string + i
    return final_string


def propertyData(row, columnlimits, worksheet_object, gt_status):
    try:
        output = {}
        output['Name'] = worksheet_object.range(row,3).value
        metrics = []
        if gt_status is True:
            for column in range(columnlimits[0],columnlimits[1]):
                if worksheet_object.range(8,column).value == "% Chg Rank" or worksheet_object.range(8,column).value == "Rank":
                    pass
                else:
                    text = str(worksheet_object.range(row,column).value)
                    text = text[:5]
                    metrics.append(text)
        else:
            for column in range(columnlimits[0],columnlimits[1]):
                if worksheet_object.range(8,column).value != "% Chg Rank":
                    metrics.append(worksheet_object.range(row,column).value)
        output['Data'] = metrics
        return output
    except Exception as e:
        print(e)
        logging.exception(e)


def grandTotalUpload(board_id, tab, grand_total_row, column_limits, worksheet_object, format):
    try:
        group_id = groupCreate(board_ident=board_id,group_name=tab)
        gt_data = propertyData(row=grand_total_row, columnlimits=column_limits, worksheet_object=worksheet_object, gt_status=True)
        print(gt_data)
        gt_name = gt_data['Name']
        gt_data = gt_data['Data']
        gt_string = stringBuilder(formatting=format,metrics=gt_data)
        newPulse(board_ident=board_id, group_ident=group_id,name_of_pulse=gt_name,data_string=gt_string)
    except Exception as e: logging.exception(e)


def propertyUpload(board_id, tab, worksheet_object, reit_limits, column_limits, abrev, format):
    try:
        group_id = groupCreate(board_ident=board_id,group_name=tab)
        if reit_limits[0] != reit_limits[1]:
            for r in range(reit_limits[0],reit_limits[1]):
                output = propertyData(row=r, columnlimits=column_limits, worksheet_object=worksheet_object, gt_status=False)
                logging.info("Output: {a}".format(a=output))
                property_name = output['Name']
                property_data = output['Data']
                for i in range(len(abrev)):
                    try:
                        property_abrev = abrev[i][property_name]
                    except:
                        pass
                property_string = stringBuilder(formatting=format, metrics=property_data)
                logging.info('string: {a}'.format(a=property_string))
                newPulse(board_ident=board_id, group_ident=group_id, name_of_pulse=property_abrev, data_string=property_string)
        else:
            output = propertyData(row=reit_limits[1], columnlimits=column_limits, worksheet_object=worksheet_object, gt_status=False)
            logging.info("Output: {a}".format(a=output))
            property_name = output['Name']
            property_data = output['Data']
            for i in range(len(abrev)):
                try:
                    property_abrev = abrev[i][property_name]
                except:
                    pass
            property_string = stringBuilder(formatting=format, metrics=property_data)
            logging.info('string: {a}'.format(a=property_string))
            newPulse(board_ident=board_id, group_ident=group_id, name_of_pulse=property_abrev,
                     data_string=property_string)
    except Exception as e: logging.exception(e)


def financialStringBuilder(formatting, data):
    format = r'\"\": \"\", '
    end_format = r'\"\": \"\"'
    final = []
    for i in range(len(formatting)):
        if i != len(formatting) - 1:
            final.append(format[:2] + formatting[i]['id'] + format[2:8] + str(data[i]) + format[8:])
        else:
            final.append(end_format[:2] + formatting[i]['id'] + end_format[2:8] + str(data[i]) + end_format[8:])
    final_string = ''
    for i in final:
        final_string = final_string + i
    return final_string


def actualData(worksheet_object, column_limits):
    l_list = ['Actual', 'Forecast', 'Budget', 'Last Year']
    acc_list = ['Room Revenue', 'Total Revenue', 'Rooms Expense', 'Total Dept Expense', 'Operating Expense',
                'House Profit', 'Fixed Expense', 'NOI B4 Interest/Other', 'NOI', 'Owner Expense', 'Net Income']
    c_index_list = []
    for i in range(4, 30):
        for j in range(column_limits[0], column_limits[1] + 1):
            if worksheet_object.range(3, j).value in l_list and worksheet_object.range(i, 1).value in acc_list:
                c_index_list.append(round(worksheet_object.range(i, j).value))
    return worksheet_object.name, c_index_list
    ...


def percentRevenuData(data_list, worksheet_object, column_limits):
    l_list = ['Actual', 'Forecast', 'Budget', 'Last Year']
    percent_list = data_list[:]
    denom_list = []
    for i in range(column_limits[0], column_limits[1] + 1):
        if worksheet_object.range(3, i).value in l_list:
            denom_list.append(round(worksheet_object.range(5, i).value))

    c = 0
    for i, v in enumerate(percent_list):
        try:
            if c < 3:
                percent_list[i] = float("{0:.1f}".format(v / denom_list[c] * 100))
                c += 1
            else:
                percent_list[i] = float("{0:.1f}".format(v / denom_list[c] * 100))
                c = 0
        except ZeroDivisionError:
            if c < 3:
                percent_list[i] = 0
                c += 1
            else:
                percent_list[i] = 0
                c = 0
    return worksheet_object.name,percent_list
    ...


def porData(data_list, worksheet_object, column_limits):
    l_list = ['Actual', 'Forecast', 'Budget', 'Last Year']
    lookup_value = 'Occupied Rooms'
    last_row = worksheet_object.range(worksheet_object.cells.last_cell.row,1).end('up').row
    print(last_row)
    # for i in range(last_row):
    #     print(worksheet_object.range(i,1).value,lookup_value,i)
    #     if worksheet_object.range(i,1).value == lookup_value:
    #         div_index = i
    #         break
    denom_list = []
    por_list = data_list[:]
    for i in range(column_limits[0], column_limits[1] + 1):
        if worksheet_object.range(3, i).value in l_list:
            denom_list.append(round(worksheet_object.range(28, i).value))
    c = 0
    for i, v in enumerate(por_list):
        try:
            if c < 3:
                por_list[i] = float("{0:.2f}".format(v / denom_list[c]))
                c += 1
            else:
                por_list[i] = float("{0:.2f}".format(v / denom_list[c]))
                c = 0
        except ZeroDivisionError:
            if c < 3:
                por_list[i] = 0
                c += 1
            else:
                por_list[i] = 0
                c = 0
    return worksheet_object.name, por_list
    ...

def limitList(worksheet_obj, last_column):
    index_list = []
    for i in range(2, last_column):
        if worksheet_obj.range(1,i).value is not None:
            index_list.append([i])
    for i in range(len(index_list)):
        if i != len(index_list)-1:
            index_list[i].append(index_list[i+1][0]-2)
        else:
            index_list[i].append(last_column-2)
    return index_list

def grandTotalDataPost(worksheet_object, format_list):
    try:
        ws = worksheet_object
        c_limits = []
        l_col = ws.range(3, ws.cells.last_cell.column).end('left').column
        for i in range(1,l_col):
            if ws.range(2, i).value == 'Grand Totals':
                c_limits.append(i)
                c_limits.append(i+8)
        key_list = []
        for i in format_list:
            for k in i:
                key_list.append(k)

        actual_group_name, act_data = actualData(worksheet_object=ws, column_limits=c_limits)
        actual_data_string = financialStringBuilder(formatting=format_list[0][key_list[0]]['columns'], data=act_data)
        actual_group_id = groupCreate(board_ident=format_list[0][key_list[0]]['id'], group_name=actual_group_name)
        newPulse(board_ident=format_list[0][key_list[0]]['id'], group_ident=actual_group_id, name_of_pulse='Grand Totals', data_string=actual_data_string)

        percent_group_name, percent_data = percentRevenuData(data_list=act_data, worksheet_object=ws, column_limits=c_limits)
        percent_data_string = financialStringBuilder(formatting=format_list[1][key_list[1]]['columns'], data=percent_data)
        percent_group_id = groupCreate(board_ident=format_list[1][key_list[1]]['id'], group_name=percent_group_name)
        newPulse(board_ident=format_list[1][key_list[1]]['id'], group_ident=percent_group_id, name_of_pulse='Grand Totals', data_string=percent_data_string)

        por_group_name, por_data = porData(data_list=act_data, worksheet_object=ws, column_limits=c_limits)
        por_data_string = financialStringBuilder(format_list[2][key_list[2]]['columns'], data=por_data)
        por_group_id = groupCreate(board_ident=format_list[2][key_list[2]]['id'], group_name=por_group_name)
        newPulse(board_ident=format_list[2][key_list[2]]['id'], group_ident=por_group_id, name_of_pulse='Grand Totals', data_string=por_data_string)

        return c_limits[0], actual_group_name
    except Exception as e: logging.exception(e)
    ...

def propertyDataPost(worksheet_object, last_column, format_list, name_of_group):
    try:
        l_list = limitList(worksheet_obj=worksheet_object,last_column=last_column)
        key_list = []
        for i in format_list:
            for k in i:
                key_list.append(k)
        actual_group_id = groupCreate(board_ident=format_list[0][key_list[0]]['id'], group_name=name_of_group)
        percent_group_id = groupCreate(board_ident=format_list[1][key_list[1]]['id'], group_name=name_of_group)
        por_group_id = groupCreate(board_ident=format_list[2][key_list[2]]['id'], group_name=name_of_group)
        for i in l_list:
            print(i)
            useless, actual_data = actualData(worksheet_object=worksheet_object,column_limits=i)
            prop_name = worksheet_object.range(2,i[0]).value
            actual_data_string = financialStringBuilder(formatting=format_list[0][key_list[0]]['columns'],data=actual_data)
            newPulse(board_ident=format_list[0][key_list[0]]['id'], group_ident=actual_group_id, name_of_pulse=prop_name,
                     data_string=actual_data_string)
            useless, percent_data = percentRevenuData(data_list=actual_data,worksheet_object=worksheet_object, column_limits=i)

            percent_data_string = financialStringBuilder(formatting=format_list[1][key_list[1]]['columns'], data=percent_data)
            newPulse(board_ident=format_list[1][key_list[1]]['id'], group_ident=percent_group_id, name_of_pulse=prop_name,
                     data_string=percent_data_string)

            useless, por_data = porData(data_list=actual_data,worksheet_object=worksheet_object, column_limits=i)
            por_data_string = financialStringBuilder(formatting=format_list[2][key_list[2]]['columns'], data=por_data)
            newPulse(board_ident=format_list[2][key_list[2]]['id'], group_ident=por_group_id, name_of_pulse=prop_name,
                     data_string=por_data_string)
    except Exception as e: logging.exception(e)

def priorNCF(board_id,ncf_format):
    try:
        query = '{boards (ids: %s) {groups {id title items{id name}}}}' % board_id
        data = {'query': query}
        r = requests.post(url=apiUrl, data=data, headers=headers)
        r = r.json()
        top_group = r['data']['boards'][0]['groups'][0]['id']
        col_list = r['data']['boards'][0]['groups'][0]['items']
        e_dict = {}
        logging.info('Columns: {a}'.format(a=col_list))
        for i in col_list:
            query2 = '{boards (ids: %s) {groups (ids: "%s"){items (ids: %s){column_values (ids: "%s"){value}}}}}' % (
            board_id, top_group, i['id'], ncf_format)
            data = {'query': query2}
            k = requests.post(url=apiUrl, data=data, headers=headers)
            k = k.json()
            k = float(k['data']['boards'][0]['groups'][0]['items'][0]['column_values'][0]['value'][1:-1])
            e_dict[i['name']] = k
        logging.info('Dictionary of values: {a}'.format(a=e_dict))
        return e_dict
    except Exception as e:logging.exception(e)

def ncfData(workbook_object,pm_ncf_dict,fund_abrev):
    try:
        account_list = ['Regular Principal (Actual)', 'Regular Interest (Actual)', 'Total Partnership Expense', 'Net Cash Burn/Flow']
        e_list = []
        abrev_list = []
        for i in range(len(fund_abrev)):
            temp = fund_abrev[i]
            for k in temp:
                abrev_list.append(temp[k])
        for i, v in enumerate(workbook_object.sheets):
            if v.name in pm_ncf_dict.keys() or v.name in abrev_list:
                e_list.append(i)
        logging.info('List of properties being uploaded: {a}'.format(a=e_list))
        data_dict = {}
        count = 0
        for i in e_list:
            ws = workbook_object.sheets[i]
            sum_index = 0
            c_index = 0
            end_row = ws.range(ws.cells.last_cell.row,1).end('up').row
            row_list = []
            for a in account_list:
                for r in range(1,end_row+1):
                    if ws.range(r,1).value == 'Summary':
                        sum_index = int(r)
                    elif ws.range(r,1).value == a:
                        row_list.append(r)
                        break
            end_column = ws.range(sum_index,ws.cells.last_cell.column).end('left').column
            for c in range(2,end_column):
                if ws.range(sum_index,c).value == "(Actual)":
                    start_column = c
                    break
            for c in range(start_column,end_column):
                if ws.range(sum_index,c).value != '(Actual)':
                    c_index = c -1
                    break
            if count == 0:
                g_index = sum_index-1
                group_name = ws.range(g_index, c_index).value
                count += 1
            data = []
            for row in row_list:
                try:
                    data.append( float("{0:.2f}".format(ws.range(row,c_index).value)))
                except TypeError:
                    data.append(0)
            data_dict[ws.name] = data
    except Exception as e: logging.exception(e)
    return data_dict, group_name

def ncfPost(board_json, workbook_object,fund_abrev):
    try:
        for key in board_json[0]:
            board_key = key
        board_id = board_json[0][board_key]['id']
        column_data = board_json[0][board_key]['column_data']
        column_id = column_data[4]['id']
        x = priorNCF(board_id=board_id, ncf_format=column_id)
        y, month_year = ncfData(workbook_object=workbook_object, pm_ncf_dict=x,fund_abrev=fund_abrev)
        group_id = groupCreate(board_ident=board_id, group_name=month_year)
        for k in y:
            try:
                y[k].append(y[k][3] + x[k])
            except KeyError:
                y[k].append(y[k][3]+0)
            d_string = financialStringBuilder(formatting=column_data, data=y[k])
            newPulse(board_ident=board_id, group_ident=group_id, name_of_pulse=k, data_string=d_string)
    except Exception as e: logging.exception(e)

def Main():
    logging.basicConfig(
        filename= 'app.log',
        level= logging.INFO,
        # format='%(levelname)s:%(asctime)s:%message)s'
    )
    logging.info('\n')
    with open('STR Board Data.json', 'r') as json_file:
        bd_fmt = json.load(json_file)
    with open('PropertyAbbreviations.JSON', 'r') as j_file:
        abrev = json.load(j_file)
    with open('Performance Board Data.JSON', 'r') as jfile:
        financial_board_format = json.load(jfile)
    with open('NCF_Board_Data.json', 'r') as infile:
        ncf_board_data = json.load(infile)
    try:
        op = options()
        file_path = browser()

        if op['Upload Type'] == '1':
            # Upload STR Data

            wb = xw.Book(file_path)
            tab_to_open = tabSelect(timeframe=op['Timeframe'],workbook_object=wb)
            ws = wb.sheets[str(tab_to_open)]
            fund_dict = data_pull(worksheet_obj=ws)
            reit_two_lims = fund_dict['LOF REIT - Fund 2']
            reit_three_lims = fund_dict['LF3 REIT - Fund 3']
            vab_lims = fund_dict['Legendary Lodging VAB QOZ']
            accel_lims = fund_dict['ACCEL II']

            if op['Fund'] == '2':
                # Check for weekly/monthly
                board_data = bd_fmt['lof2']
                gt_row = reit_two_lims[1]
                fund_abrev = abrev['lof2']
                if op['Timeframe'] == '2':  # Fund 2, Monthly Time Frame
                    logging.info('Uploading Monthly Data for LOF 2')
                    # Setting Grand Total variables
                    gt_board_id = board_data[0]["LOF2 Monthly STR - Grand Total"]['id']
                    gt_format = board_data[0]["LOF2 Monthly STR - Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    # Setting Property variables
                    prop_board_id = board_data[1]["LOF2 Monthly STR - Properties"]['id']
                    prop_format = board_data[1]["LOF2 Monthly STR - Properties"]['column_data']
                    # Grand Total Upload
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=reit_two_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
                elif op['Timeframe'] == '1':
                    logging.info('Uploading Weekly Data for LOF 2')
                    # Setting Grand Total variables
                    gt_board_id = board_data[2]["LOF2 Weekly STR - Grand Total"]['id']
                    gt_format = board_data[2]["LOF2 Weekly STR - Grand Total"]['column_data']
                    # gt_format = gt_format[:-1]
                    # Setting Property variables
                    prop_board_id = board_data[3]["LOF2 Weekly STR - Properties"]['id']
                    prop_format = board_data[3]["LOF2 Weekly STR - Properties"]['column_data']
                    # Grand Total Upload
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=reit_two_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
                # elif op['Timeframe'] == '1':
            elif op['Fund'] == '3':
                board_data = bd_fmt['lf3']
                gt_row = reit_three_lims[1]
                fund_abrev = abrev['lf3']
                if op['Timeframe'] == '2':  # Fund 3, Monthly Time Frame
                    # Setting Grand Total variables
                    logging.info('Uploading Monthly Data for LF 3')
                    gt_board_id = board_data[0]["LF3 Monthly STR - Grand Total"]['id']
                    gt_format = board_data[0]["LF3 Monthly STR - Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    # Setting Property variables
                    prop_board_id = board_data[1]["LF3 Monthly STR - Properties"]['id']
                    prop_format = board_data[1]["LF3 Monthly STR - Properties"]['column_data']
                    # Grand Total Upload
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=reit_three_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
                elif op['Timeframe'] == '1':
                    logging.info('Uploading Weekly Data for LF 3')
                    # Setting Grand Total variables
                    gt_board_id = board_data[2]["LF3 Weekly STR - Grand Total"]['id']
                    gt_format = board_data[2]["LF3 Weekly STR - Grand Total"]['column_data']

                    # Setting Property variables
                    prop_board_id = board_data[3]["LF3 Weekly STR - Properties"]['id']
                    prop_format = board_data[3]["LF3 Weekly STR - Properties"]['column_data']
                    # Grand Total Upload
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=reit_three_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
            elif op['Fund'] == '4':
                board_data = bd_fmt['Accel 2']
                gt_row = accel_lims[1]
                fund_abrev = abrev['Accel 2']
                if op['Timeframe'] == '2':
                    logging.info('Uploading Monthly Data for Accel 2')
                    gt_board_id = board_data[3]["Accel II Monthly STR Grand Total"]['id']
                    gt_format = board_data[3]["Accel II Monthly STR Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    # Setting Property variables
                    prop_board_id = board_data[2]["Accel II Monthly STR-Properties"]['id']
                    prop_format = board_data[2]["Accel II Monthly STR-Properties"]['column_data']
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=accel_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
                elif op['Timeframe'] == '1':
                    logging.info('Uploading Weekly Data for Accel 2')
                    gt_board_id = board_data[1]["Accel II Weekly STR-Grand Total"]['id']
                    gt_format = board_data[1]["Accel II Weekly STR-Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    # Setting Property variables
                    prop_board_id = board_data[0]["Accel II Weekly STR-Properties"]['id']
                    prop_format = board_data[0]["Accel II Weekly STR-Properties"]['column_data']
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=accel_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
            elif op['Fund'] == '5':
                board_data = bd_fmt['VABQOZ']
                fund_abrev = abrev['VAB QOZ']
                if op['Timeframe'] =='2':
                    board_id = board_data[0]["VABQOZ Monthly STR Properties"]['id']
                    fmat = board_data[0]["VABQOZ Monthly STR Properties"]['column_data']
                    propertyUpload(board_id=board_id, tab=tab_to_open, reit_limits=vab_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=fmat)
                elif op['Timeframe'] == '1':
                    board_id = board_data[1]["VABQOZ Weekly STR"]['id']
                    fmat = board_data[1]["VABQOZ Weekly STR"]['column_data']
                    propertyUpload(board_id=board_id, tab=tab_to_open, reit_limits=vab_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=fmat)
            elif op['Fund'] == '1':
                if op['Timeframe'] == '2':
                    board_data = bd_fmt['lof2']
                    gt_row = reit_two_lims[1]
                    gt_board_id = board_data[0]["LOF2 Monthly STR - Grand Total"]['id']
                    gt_format = board_data[0]["LOF2 Monthly STR - Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    fund_abrev = abrev['lof2']
                    # Setting Property variables
                    prop_board_id = board_data[1]["LOF2 Monthly STR - Properties"]['id']
                    prop_format = board_data[1]["LOF2 Monthly STR - Properties"]['column_data']
                    # Grand Total Upload
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=reit_two_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
                    board_data = bd_fmt['lf3']
                    gt_row = reit_three_lims[1]
                    gt_board_id = board_data[0]["LF3 Monthly STR - Grand Total"]['id']
                    gt_format = board_data[0]["LF3 Monthly STR - Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    fund_abrev = abrev['lf3']
                    # Setting Property variables
                    prop_board_id = board_data[1]["LF3 Monthly STR - Properties"]['id']
                    prop_format = board_data[1]["LF3 Monthly STR - Properties"]['column_data']
                    # Grand Total Upload
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=reit_three_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
                    board_data = bd_fmt['Accel 2']
                    gt_row = accel_lims[1]
                    fund_abrev = abrev['Accel 2']
                    gt_board_id = board_data[3]["Accel II Monthly STR Grand Total"]['id']
                    gt_format = board_data[3]["Accel II Monthly STR Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    # Setting Property variables
                    prop_board_id = board_data[2]["Accel II Monthly STR-Properties"]['id']
                    prop_format = board_data[2]["Accel II Monthly STR-Properties"]['column_data']
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=accel_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)

                    board_data = bd_fmt['VABQOZ']
                    gt_row = vab_lims[1]
                    fund_abrev = abrev['VAB QOZ']

                    # Setting Property variables
                    prop_board_id = board_data[0]["VABQOZ Monthly STR Properties"]['id']
                    prop_format = board_data[0]["VABQOZ Monthly STR Properties"]['column_data']

                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=accel_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev,
                                   format=prop_format)
                elif op['Timeframe'] == '1':
                    board_data = bd_fmt['lof2']
                    gt_row = reit_two_lims[1]
                    # Setting Grand Total variables
                    gt_board_id = board_data[2]["LOF2 Weekly STR - Grand Total"]['id']
                    gt_format = board_data[2]["LOF2 Weekly STR - Grand Total"]['column_data']
                    fund_abrev = abrev['lof2']
                    # Setting Property variables
                    prop_board_id = board_data[3]["LOF2 Weekly STR - Properties"]['id']
                    prop_format = board_data[3]["LOF2 Weekly STR - Properties"]['column_data']
                    # Grand Total Upload
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims,
                                     worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=reit_two_lims,
                                   column_limits=column_lims,
                                   worksheet_object=ws, abrev=fund_abrev, format=prop_format)
                    board_data = bd_fmt['lf3']
                    fund_abrev = abrev['lf3']
                    gt_row = reit_three_lims[1]
                    # Setting Grand Total variables
                    gt_board_id = board_data[2]["LF3 Weekly STR - Grand Total"]['id']
                    gt_format = board_data[2]["LF3 Weekly STR - Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    # Setting Property variables
                    prop_board_id = board_data[3]["LF3 Weekly STR - Properties"]['id']
                    prop_format = board_data[3]["LF3 Weekly STR - Properties"]['column_data']
                    # Grand Total Upload
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=reit_three_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
                    gt_board_id = board_data[1]["Accel II Weekly STR-Grand Total"]['id']
                    gt_format = board_data[1]["Accel II Weekly STR-Grand Total"]['column_data']
                    gt_format = gt_format[:-1]
                    # Setting Property variables
                    prop_board_id = board_data[0]["Accel II Weekly STR-Properties"]['id']
                    prop_format = board_data[0]["Accel II Weekly STR-Properties"]['column_data']
                    grandTotalUpload(board_id=gt_board_id, tab=tab_to_open, grand_total_row=gt_row,
                                     column_limits=column_lims, worksheet_object=ws, format=gt_format)
                    # Properties Upload
                    propertyUpload(board_id=prop_board_id, tab=tab_to_open, reit_limits=accel_lims,
                                   column_limits=column_lims, worksheet_object=ws, abrev=fund_abrev, format=prop_format)
        elif op['Upload Type'] == '2':
            # Upload Financial Data
            wb = xw.Book(file_path)
            tab_to_open = tabSelect(timeframe='0',workbook_object=wb)
            ws = wb.sheets[str(tab_to_open)]
            if op['Fund'] == '2':
                logging.info('Uploading Financial Data for LOF 2')
                gt_board_format = financial_board_format['LOF2'][:3]
                property_board_format = financial_board_format['LOF2'][3:]
                limit, group_name = grandTotalDataPost(worksheet_object=ws, format_list=gt_board_format)
                propertyDataPost(worksheet_object=ws,last_column=limit,format_list=property_board_format,name_of_group=group_name)
                ...
            elif op['Fund'] == '3':
                logging.info('Uploading Financial Data for LF 3')
                gt_board_format = financial_board_format['LF3'][:3]
                property_board_format = financial_board_format['LF3'][3:]
                limit, group_name = grandTotalDataPost(worksheet_object=ws, format_list=gt_board_format)
                propertyDataPost(worksheet_object=ws, last_column=limit, format_list=property_board_format,
                                 name_of_group=group_name)
            elif op['Fund'] == '4':
                c_lims = [2,12]
                actual_group_name, act_data = actualData(worksheet_object=ws,column_limits=c_lims)
                property_board_format = financial_board_format['Accel II'][:3]
                propertyDataPost(worksheet_object=ws, last_column=c_lims[1], format_list=property_board_format,name_of_group=actual_group_name)
                # Accel II
            elif op['Fund'] == '5':
                # VAB QOZ
                c_lims = [2,12]
                actual_group_name, act_data = actualData(worksheet_object=ws, column_limits=c_lims)
                property_board_format = financial_board_format['VABQOZ'][:3]
                propertyDataPost(worksheet_object=ws, last_column=c_lims[1], format_list=property_board_format,name_of_group=actual_group_name)
                ...
        elif op['Upload Type'] == '3':
            if op['Fund'] == '2':
                abrev = abrev['lof2']
                logging.info("Uploading NCF Data for LOF2")
                ncf_board_data = ncf_board_data['LOF2']
                wb = xw.Book(file_path)
                ncfPost(board_json=ncf_board_data, workbook_object=wb, fund_abrev=abrev)
                ...
            elif op['Fund'] == '3':
                abrev = abrev['lf3']
                ncf_board_data= ncf_board_data['LF3']
                wb = xw.Book(file_path)
                ncfPost(board_json=ncf_board_data, workbook_object=wb, fund_abrev=abrev)
                ...

    except Exception as e: logging.exception(e)
    Main()

Main()