import pandas as pd
import numpy as np
import openpyxl as opxl
import os.path
import pycountry
import pycountry_convert as pc
import datetime as dt
import tkinter as tk
from tkinter import filedialog
from FillFuncs import fillVWR_Old, fillThomas_Old, fillFisher_Old, fillFisher_Enrichment, fillVWR_Enrichment, fillVWR_New, fillThomas_New, fillFisher_New, VWREnrichmentDriver, fillGlobalProductRevision, fillGlobalProductRevisionChemicals


def importExcelSheets():
    lbl_main['text'] = "Importing May Magento..."
    root.update()
    magento = pd.read_excel('database_sheets/magento_may.xlsx')
    lbl_main['text'] = "Importing July Magento..."
    root.update()
    new_magento = pd.read_excel('database_sheets/magento_july.xlsx')
    lbl_main['text'] = "Importing September Magento"
    root.update()
    magento_sept = pd.read_excel('database_sheets/magento_sept.xlsx')
    lbl_main['text'] = "Importing Lot Master..."
    root.update()
    lot_master = pd.read_excel('database_sheets/lot_master.xlsx', dtype = object)
    lbl_main['text'] = "Importing PRMS..."
    root.update()
    prms = pd.read_excel('database_sheets/prms.xlsx')
    lbl_main['text'] = "Importing UNSPSC Codes..."
    root.update()
    unspsc_codes = pd.read_excel('database_sheets/unspsc_codes.xlsx')
    unspsc_codes.columns = unspsc_codes.iloc[0]
    lbl_main['text'] = "Importing country of origin info..."
    root.update()
    origin = pd.read_excel('database_sheets/country_of_origin.xlsx', dtype=object)
    pd.set_option("max_rows", None)
    pd.set_option("max_columns", None)
    return [magento, new_magento, lot_master, prms, unspsc_codes, origin, magento_sept]

def fillAll_Old_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    database_sheets = importExcelSheets()
    lbl_main['text'] = 'Filling VWR...'
    root.update()
    fillVWR_Old(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[4], database_sheets[5])
    lbl_main['text'] = 'Filling Thomas...'
    root.update()
    fillThomas_Old(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[4], database_sheets[5], database_sheets[6])
    lbl_main['text'] = 'Filling Fisher...'
    root.update()
    fillFisher_Old(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[4], database_sheets[5], database_sheets[6])
    lbl_main['text'] = 'All Done...'
    root.update()

def fillVWR_Old_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    database_sheets = importExcelSheets()
    lbl_main['text'] = 'Filling VWR...'
    root.update()
    fillVWR_Old(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[4], database_sheets[5])
    lbl_main['text'] = 'All Done...'
    root.update()

def fillFisher_Old_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    database_sheets = importExcelSheets()
    lbl_main['text'] = 'Filling Fisher...'
    root.update()
    fillFisher_Old(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[4], database_sheets[5], database_sheets[6])
    lbl_main['text'] = 'All Done...'
    root.update()

def fillThomas_Old_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    database_sheets = importExcelSheets()
    lbl_main['text'] = 'Filling Thomas...'
    root.update()
    fillThomas_Old(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[4], database_sheets[5], database_sheets[6])
    lbl_main['text'] = 'All Done...'
    root.update()

def fisherEnrichmentHelper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    database_sheets = importExcelSheets()
    images = pd.read_excel('database_sheets/MPBIO_Products_Images_20_10_21.xlsx')
    lbl_main['text'] = 'Filling Fisher Enrichment Form (This one may take a while)'
    root.update()
    fillFisher_Enrichment(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[6], database_sheets[4], database_sheets[5], images)
    lbl_main['text'] = 'Done'

def VWREnrichmentHelper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    lbl_main['text'] = "Importing Old Magento..."
    root.update()
    magento = pd.read_excel('database_sheets/magento_may.xlsx')
    lbl_main['text'] = 'Importing Sept Magento...'
    root.update()
    magento_sept = pd.read_excel('database_sheets/magento_sept.xlsx')
    lbl_main['text'] = 'Importing PRMS...'
    root.update()
    prms = pd.read_excel('database_sheets/prms.xlsx')
    lbl_main['text'] = 'Importing Product Categories...'
    root.update()
    categories = pd.read_excel('database_sheets/product_categories.xlsx')
    lbl_main['text'] = 'Filling VWR Enrichment Forms'
    root.update()
    VWREnrichmentDriver(filename, magento_sept, prms, magento, categories)
    lbl_main['text'] = 'Done'

def packEnrichmentButtons():
    bt_fisher_enrichment.pack(side=tk.TOP)
    bt_VWR_enrichment.pack(side=tk.TOP)

def importNewProductAdd(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    lbl_main['text'] = 'Importing New Product Add Form...'
    data = pd.ExcelFile(filename)
    product_manager = pd.read_excel(data, 'Product Manager', dtype=object)
    prms2 = pd.read_excel(data, 'PRMS', dtype=object)
    e_marketing = pd.read_excel(data, 'eMarketing', dtype=object)
    return [product_manager, prms2, e_marketing]

def fillVWR_New_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    lbl_main['text'] = 'Importing Form...'
    root.update()
    database_sheets = importNewProductAdd(filename)
    lbl_main['text'] = 'Filling VWR...'
    root.update()
    fillVWR_New(database_sheets[0], database_sheets[1], database_sheets[2])
    lbl_main['text'] = 'Done'

def fillThomas_New_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    lbl_main['text'] = 'Importing Form...'
    root.update()
    database_sheets = importNewProductAdd(filename)
    lbl_main['text'] = 'Filling Thomas...'
    root.update()
    fillThomas_New(database_sheets[0], database_sheets[1], database_sheets[2])
    lbl_main['text'] = 'Done'

def fillFisher_New_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    lbl_main['text'] = 'Importing Form...'
    root.update()
    database_sheets = importNewProductAdd(filename)
    lbl_main['text'] = 'Filling Fisher...'
    root.update()
    fillFisher_New(database_sheets[0], database_sheets[1], database_sheets[2])
    lbl_main['text'] = 'Done'

def fillAll_New_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    lbl_main['text'] = 'Importing Form...'
    root.update()
    database_sheets = importNewProductAdd(filename)
    lbl_main['text'] = 'Filling Fisher...'
    root.update()
    fillFisher_New(database_sheets[0], database_sheets[1], database_sheets[2])
    lbl_main['text'] = 'Filling VWR...'
    root.update()
    fillVWR_New(database_sheets[0], database_sheets[1], database_sheets[2])
    lbl_main['text'] = 'Filling Thomas...'
    root.update()
    fillThomas_New(database_sheets[0], database_sheets[1], database_sheets[2])
    lbl_main['text'] = 'Done'

def fillRevision_Normal_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    database_sheets = importExcelSheets()
    lbl_main['text'] = 'Filling Global Product Revision...'
    root.update()
    fillGlobalProductRevision(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[4], database_sheets[5], database_sheets[6])
    lbl_main['text'] = 'Done'

def fillRevision_Chemicals_Helper(filename):
    if filename == '':
        lbl_main['text'] = 'File not found... Please select a valid file'
        return
    database_sheets = importExcelSheets()
    lbl_main['text'] = 'Filling Global Product Revision Chemicals...'
    root.update()
    fillGlobalProductRevisionChemicals(filename, database_sheets[0], database_sheets[1], database_sheets[2], database_sheets[3], database_sheets[4], database_sheets[5], database_sheets[6])
    lbl_main['text'] = 'Done'

def packNewButtons():
    bt_VWR_new.pack(side=tk.TOP)
    bt_Fisher_new.pack(side=tk.TOP)
    bt_Thomas_new.pack(side=tk.TOP)
    bt_all_new.pack(side=tk.TOP)

def packOldButtons():
    bt_VWR_old.pack(side=tk.TOP)
    bt_Fisher_old.pack(side=tk.TOP)
    bt_Thomas_old.pack(side=tk.TOP)
    bt_all_old.pack(side=tk.TOP)

def packRevisionButtons():
    bt_fill_revision_normal.pack(side=tk.LEFT)
    bt_fill_revision_chemicals.pack(side=tk.RIGHT)

root = tk.Tk()

canvas = tk.Canvas(root)
canvas.grid(columnspan=4, rowspan=2)

frame_label = tk.Frame(root)
frame_buttons = tk.Frame(root)

frame_buttons_middle_left = tk.Frame(frame_buttons)
frame_buttons_middle_right = tk.Frame(frame_buttons)
frame_buttons_left = tk.Frame(frame_buttons)
frame_buttons_right = tk.Frame(frame_buttons)

frame_old = tk.Frame(frame_buttons_left)
frame_old_top = tk.Frame(frame_old)
frame_old_bottom = tk.Frame(frame_old)
frame_old_VWR = tk.Frame(frame_old_bottom)
frame_old_Fisher = tk.Frame(frame_old_bottom)
frame_old_Thomas = tk.Frame(frame_old_bottom)
frame_old_all = tk.Frame(frame_old_bottom)

frame_enrichment = tk.Frame(frame_buttons_right)
frame_new = tk.Frame(frame_buttons_middle_left)
frame_enrichment_top = tk.Frame(frame_enrichment)
frame_enrichment_bottom = tk.Frame(frame_enrichment)
frame_enrichment_fisher = tk.Frame(frame_enrichment_bottom)
frame_enrichment_VWR = tk.Frame(frame_enrichment_bottom)

frame_revision = tk.Frame(frame_buttons_middle_right)
frame_revision_top = tk.Frame(frame_revision)
frame_revision_bottom = tk.Frame(frame_revision)

frame_label.grid(row=0, columnspan=4)
frame_buttons.grid(row=1, columnspan=4, rowspan=1)
frame_buttons_left.grid(column=0, row=1)
frame_buttons_middle_left.grid(column=1, row=1)
frame_buttons_middle_right.grid(column=2, row=1)
frame_buttons_right.grid(column=3, row=1)

frame_revision_top.pack(side=tk.TOP)
frame_revision_bottom.pack(side=tk.BOTTOM)

frame_old.pack()
frame_old_top.pack(side=tk.TOP)
frame_old_bottom.pack(side=tk.BOTTOM)
frame_old_VWR.pack(side=tk.LEFT)
frame_old_Fisher.pack(side=tk.LEFT)
frame_old_Thomas.pack(side=tk.LEFT)
frame_old_all.pack(side=tk.LEFT)

frame_enrichment.pack()
frame_new.pack()
frame_revision.pack()
frame_enrichment_bottom.pack(side=tk.BOTTOM)
frame_enrichment_top.pack(side=tk.TOP)
frame_enrichment_fisher.pack(side=tk.LEFT)
frame_enrichment_VWR.pack(side=tk.RIGHT)

frame_new_top = tk.Frame(frame_new)
frame_new_bottom = tk.Frame(frame_new)
frame_new_VWR = tk.Frame(frame_new_bottom)
frame_new_Fisher = tk.Frame(frame_new_bottom)
frame_new_Thomas = tk.Frame(frame_new_bottom)
frame_new_all = tk.Frame(frame_new_bottom)
frame_new_top.pack(side=tk.TOP)
frame_new_bottom.pack(side=tk.BOTTOM)
frame_new_VWR.pack(side=tk.LEFT)
frame_new_Fisher.pack(side=tk.LEFT)
frame_new_Thomas.pack(side=tk.LEFT)
frame_new_all.pack(side=tk.LEFT)

lbl_main = tk.Label(frame_label, text='Select Option')
lbl_main.pack()

bt_VWR_old = tk.Button(frame_old_VWR, text='Fill VWR', command=lambda: bt_VWR_old_askFile.pack(side=tk.BOTTOM))
bt_Fisher_old = tk.Button(frame_old_Fisher, text='Fill Fisher', command=lambda: bt_Fisher_old_askFile.pack(side=tk.BOTTOM))
bt_Thomas_old = tk.Button(frame_old_Thomas, text='Fill Thomas', command=lambda: bt_Thomas_old_askFile.pack(side=tk.BOTTOM))
bt_all_old = tk.Button(frame_old_all, text='Fill All Three', command=lambda: bt_all_old_askFile.pack(side=tk.BOTTOM))
bt_VWR_old_askFile = tk.Button(frame_old_VWR, text='Choose Product SKUs', command=lambda: fillVWR_Old_Helper(tk.filedialog.askopenfilename()))
bt_Fisher_old_askFile = tk.Button(frame_old_Fisher, text='Choose Product SKUs', command=lambda: fillFisher_Old_Helper(tk.filedialog.askopenfilename()))
bt_Thomas_old_askFile = tk.Button(frame_old_Thomas, text='Choose Product SKUs', command=lambda: fillThomas_Old_Helper(tk.filedialog.askopenfilename()))
bt_all_old_askFile = tk.Button(frame_old_all, text='Choose Product SKUs', command=lambda: fillAll_Old_Helper(tk.filedialog.askopenfilename()))
# bt_askFile_old = tk.Button(frame_old_bottom, text='Select SKU File', command=lambda: fillOld(tk.filedialog.askopenfilename()))
bt_fill_old = tk.Button(frame_old_top, text='Current Database Output', command=lambda: packOldButtons())
# bt_askFile_old.pack_forget()
bt_fill_old.pack(side=tk.TOP)
bt_VWR_old.pack_forget()
bt_Fisher_old.pack_forget()
bt_Thomas_old.pack_forget()
bt_all_old.pack_forget()
bt_VWR_old_askFile.pack_forget()
bt_Fisher_old_askFile.pack_forget()
bt_Thomas_old_askFile.pack_forget()
bt_all_old_askFile.pack_forget()

bt_VWR_new_askFile = tk.Button(frame_new_VWR, text='Choose New Product Add Form', command=lambda: fillVWR_New_Helper(tk.filedialog.askopenfilename()))
bt_Fisher_new_askFile = tk.Button(frame_new_Fisher, text='Choose New Product Add Form', command=lambda: fillFisher_New_Helper(tk.filedialog.askopenfilename()))
bt_Thomas_new_askFile = tk.Button(frame_new_Thomas, text='Choose New Product Add Form', command=lambda: fillThomas_New_Helper(tk.filedialog.askopenfilename()))
bt_all_new_askFile = tk.Button(frame_new_all, text='Choose New Product Add Form', command=lambda: fillAll_New_Helper(tk.filedialog.askopenfilename()))
bt_VWR_new = tk.Button(frame_new_VWR, text='Fill VWR', command=lambda: bt_VWR_new_askFile.pack(side=tk.BOTTOM))
bt_Fisher_new = tk.Button(frame_new_Fisher, text='Fill Fisher', command=lambda: bt_Fisher_new_askFile.pack(side=tk.BOTTOM))
bt_Thomas_new = tk.Button(frame_new_Thomas, text='Fill Thomas', command=lambda: bt_Thomas_new_askFile.pack(side=tk.BOTTOM))
bt_all_new = tk.Button(frame_new_all, text='Fill All Forms', command=lambda: bt_all_new_askFile.pack(side=tk.BOTTOM))
bt_fill_new = tk.Button(frame_new_top, text='Output Using a New Product Add Form', command=lambda: packNewButtons())
bt_fill_new.pack(side=tk.TOP)
bt_VWR_new.pack_forget()
bt_Fisher_new.pack_forget()
bt_Thomas_new.pack_forget()
bt_all_new.pack_forget()
bt_VWR_new_askFile.pack_forget()
bt_Fisher_new_askFile.pack_forget()
bt_Thomas_new_askFile.pack_forget()
bt_all_new_askFile.pack_forget()

bt_askFile_fisher_enrichment = tk.Button(frame_enrichment_fisher, text='Select Fisher Enrichment File', command=lambda: fisherEnrichmentHelper(tk.filedialog.askopenfilename()))
bt_askFile_VWR_enrichment = tk.Button(frame_enrichment_VWR, text='Select SKU File for VWR Enrichment', command=lambda: VWREnrichmentHelper(tk.filedialog.askopenfilename()))
bt_fisher_enrichment = tk.Button(frame_enrichment_fisher, text='Fisher Enrichment', command=lambda: bt_askFile_fisher_enrichment.pack(side=tk.BOTTOM))
bt_VWR_enrichment = tk.Button(frame_enrichment_VWR, text='VWR Enrichment', command=lambda: bt_askFile_VWR_enrichment.pack(side=tk.BOTTOM))
bt_fill_enrichment = tk.Button(frame_enrichment_top, text='Enrichment Ouputs', command=packEnrichmentButtons)
bt_fisher_enrichment.pack_forget()
bt_VWR_enrichment.pack_forget()
bt_fill_enrichment.pack()

bt_fill_revision = tk.Button(frame_revision_top, text='Fill Revision', command=packRevisionButtons)
bt_fill_revision_normal = tk.Button(frame_revision_bottom, text='Fill Basic Revision', command=lambda: fillRevision_Normal_Helper(tk.filedialog.askopenfilename()))
bt_fill_revision_chemicals = tk.Button(frame_revision_bottom, text='Fill Revision Chemicals', command=lambda: fillRevision_Chemicals_Helper(tk.filedialog.askopenfilename()))
bt_fill_revision.pack()
bt_fill_revision_chemicals.pack_forget()
bt_fill_revision_normal.pack_forget()

root.minsize(300, 40)
root.title('Autofill')
root.mainloop()