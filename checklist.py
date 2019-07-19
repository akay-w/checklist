# -*- coding: utf-8 -*-
"""
Created on Mon May 27 14:24:18 2019

@author: a-whalen
"""

import os
from editpyxl import Workbook

def date_format(date):
    if len(date) < 4:
        return date
    month = date[:2]
    day = date[2:]
    return month + "/" + day

sagyou = os.path.normpath(r"") #insert path here
for root, folders, files in os.walk(sagyou):
    for file in files:
        if file == r"checklist.xlsx":
            checklist = os.path.join(os.path.normpath(root), file)
            split_path = checklist.split("\\")
            folder_name = split_path[-3]
            split_folder_name = folder_name.split("_")
            order_number = split_folder_name[0]
            client = split_folder_name[1]
            try:
                langs = split_folder_name[2]
            except IndexError:
                langs = split_folder_name
            if "to" in langs:
                split_lang = langs.split("to")
                src = split_lang[0]
                tar = split_lang[1]
            elif "-" in langs:
                split_lang = langs.split("-")
                src = split_lang[0]
                tar = split_lang[1]
            else:
                if langs.isalpha():
                    src = langs[:3]
                    tar = langs[3:]
                else:
                    src = langs[:1]
                    tar = langs[1:]
            base = checklist.split("\\7_Checklist")[0]
            client_fldr_date = None
            client_path = os.path.join(base,"2_Client")
            for client_root, client_fldr, client_file in os.walk(client_path):
                for c in client_fldr:
                    if "XXXX" not in c and not client_fldr_date:
                        client_fldr_date = date_format(c)
                break
            trans_path = os.path.join(base, "3_Translation")
            trans_fldr_date, outsource_date, outsource_deliv_date = None, None, None
            tm_mente_date = ""
            for trans_root, trans_fldr, trans_file in os.walk(trans_path):
                for t in trans_fldr:
                    if "XXXX" not in t:
                        if "Translation" in t and not trans_fldr_date:
                            trans_fldr_date = date_format(t.split("_")[0])
                        if "Outsource_Order" in t and not outsource_date:
                            outsource_date = date_format(t.split("_")[0])
                        if "Outsource_Delivery" in t and not outsource_deliv_date:
                            outsource_deliv_date = date_format(t.split("_")[0])
            check_path = os.path.join(base, "4_Translation_Check")
            dtp_path = os.path.join(base, "5_DTP")
            dtp_bool = False
            for dtp_root, dtp_fldr, dtp_file in os.walk(dtp_path):
                for d in dtp_fldr:
                    print(d)
                    if "XXXX" not in d and "DTP" in d:
                        print("yes")
                        dtp_bool = True            
            deliv_path = os.path.join(base, "6_Delivery")
            deliv_date = None
            for deliv_root, deliv_fldr, deliv_file in os.walk(deliv_path):
                for d in deliv_fldr:
                    if "XXXX" not in d:
                        deliv_date = date_format(d)
            tm_mente_path = os.path.join(base,"8_Other")
            for tm_root, tm_fldr, tm_file in os.walk(tm_mente_path):
                for t in tm_fldr:
                    if "XXXX" not in t and "TM_Mente" in t:
                        tm_mente_date = date_format(t.split("_")[0])
            wb = Workbook()
            wb.open(checklist)
            ws = wb.active
            ws.cell("B3").value = order_number
            ws.cell("B4").value = client
            ws.cell("B5").value = src
            ws.cell("D5").value = tar
            if client_fldr_date:
                ws.cell("H8").value = client_fldr_date
                ws.cell("H11").value = client_fldr_date
                ws.cell("H13").value = client_fldr_date
                ws.cell("H16").value = client_fldr_date
            for i in range(8,16):
                ws.cell(row = i, column = 7).value = "〇"
            for i in range(19,27):
                ws.cell(row = i, column = 7).value = "〇"
            for i in range(33,35):
                ws.cell(row = i, column = 7).value = "〇"
            if outsource_date:
                ws.cell("G16").value = "〇"
                ws.cell("H16").value = outsource_date
            if tm_mente_date:
                ws.cell("G32").value = "〇"
                ws.cell("H32").value = tm_mente_date
            if deliv_date:
                if dtp_bool == True:
                    ws.cell("G29").value = "〇"
                    ws.cell("H29").value = deliv_date
                for i in range(30,32):
                    ws.cell(row = i, column = 7).value = "〇"
                    ws.cell(row = i, column = 8).value = deliv_date
                for i in range(33,35):
                    ws.cell(row = i, column = 7).value = "〇"
                    ws.cell(row = i, column = 8).value = deliv_date
                if outsource_deliv_date:
                    ws.cell("H19").value = outsource_deliv_date
                elif trans_fldr_date:
                   ws.cell("H19").value = trans_fldr_date
                ws.cell("H23").value = deliv_date
                
            newFilename = 'Jidou_' + os.path.basename(checklist)
            wb.save(os.path.join(os.path.dirname(checklist), newFilename))
            wb.close()



