from tkinter import *
from tkinter import ttk
from tkinter.ttk import *
import pandas as pd
import numpy as np
import time
import os
import xlwt
from xlwt import Workbook
import easygui as eg

#########################################################################################
#########################################################################################
##                                                                                     ##
##                               #### HOW TO IMPORT ####                               ##
##                                                                                     ##
## home desktop | import tkinter_lib                                                   ##
##                                                                                     ##
## import sys                                                                          ##
## sys.path.append('C:\\Users\\joo09\\Documents\\GitHub\\tkinter_gui')                 ##
## import tkinter_lib as lib                                                           ##
##                                                                                     ##
##                                                                                     ##
##                                                                                     ##
##                           #### HOW TO RE - IMPORT ####                              ##
##                                                                                     ##
## from imp import reload                                                              ##
## reload(lib)                                                                         ##
##                                                                                     ##
#########################################################################################
#########################################################################################

# save location
# saveloc = 'C:\\Users\\joo09\\Documents\\GitHub\\tkinter_gui\\saved_excel'


#########################################################################################
######################################## < theme > ######################################


# root.tk.call("source", "sun-valley.tcl")
# root.tk.call("set_theme", "light")

# def change_theme():
#     # NOTE: The theme's real name is sun-valley-<mode>
#     if root.tk.call("ttk::style", "theme", "use") == "sun-valley-dark":
#         # Set light theme
#         root.tk.call("set_theme", "light")
#     else:
#         # Set dark theme
#         root.tk.call("set_theme", "dark")

# # Remember, you have to use ttk widgets
# button = ttk.Button(big_frame, text="Change theme!", command=change_theme)
# button.pack()


#########################################################################################
###################################### < excel read > ###################################

def read_excel(excel) :
    df = pd.read_excel(excel)
    df = delun(df)

    return df


def delun(df) :
    if 'Unnamed: 0' in df.columns :
        df.drop('Unnamed: 0', axis = 1, inplace = True)
    if 'Unnamed: 0.1' in df.columns :
        df.drop('Unnamed: 0.1', axis = 1, inplace = True)

    return df
#########################################################################################
######################################## < button > #####################################

# def savebutton() :
#     savebutton = Button(root, command = fastsave(), text = 'save')
    


#########################################################################################
######################################### < save > ######################################

def saveas(saveloc):
    documento=Workbook()
    insert_celdas = documento.add_sheet('Nomina')
    cambio=insert_celdas.set_portrait(False)
    extension =["*.xlsx"]
    gplanilla=eg.filesavebox(msg="save file",
        title="",default='{}'.format(time.strftime('%Y_%m_%d', time.localtime(time.time()))),
        filetypes=extension) 
    if (gplanilla is not None) and (len(gplanilla)!=0):
        documento.save(str(gplanilla)+".xlsx")
        print("canceled the process of generation")
    else:
        print("generation completed")
        
        return False
    
    lastfile = '{}.xlsx'.format(time.strftime('%Y_%m_%d', time.localtime(time.time())))
    last_file = pd.DataFrame(columns = ['name'])
    last_file.loc[0, 'name'] = lastfile
    os.chdir(saveloc)
    last_file.to_excel('lastfile.xlsx')


def fastsave(df, saveloc) :
    os.chdir(saveloc)
#     os.remove(lastfile)
    df.to_excel('{}.xlsx'.format(time.strftime('%Y_%m_%d', time.localtime(time.time()))))
    lastfile = '{}.xlsx'.format(time.strftime('%Y_%m_%d', time.localtime(time.time())))
    last_file = pd.DataFrame(columns = ['name'])
    last_file.loc[0, 'name'] = lastfile
    os.chdir(saveloc)
    last_file.to_excel('lastfile.xlsx')

#########################################################################################
######################################### < init > ######################################

def init_load(saveloc) :
    os.chdir(saveloc)
    if 'lastfile.xlsx' in os.listdir(saveloc) :
        lastdf = read_excel('lastfile.xlsx')
        lastname = lastdf.loc[0, 'name']
        if lastname in os.listdir(saveloc) :
            df = read_excel(lastname)
        else :
            df = read_excel('sample.xlsx')
        
    else :
        df = read_excel('sample.xlsx')

    return df

def init_tree(tree, df, root) :
    
    tree = ttk.Treeview(root, columns = ['number'] + df.columns, displaycolumns = ['number'] + df.columns)
    tree.pack()
    
    for i in range(len(df.columns) + 1) :
        if i == 0 :
            tree.column('#{}'.format(i), width = 50, anchor = 'center')
            tree.heading('#{}'.format(i), text = 'num', anchor = 'center')
        elif i == len(df.columns) :
            tree.column('#{}'.format(i), width = 300, anchor = 'w')
            tree.heading('#{}'.format(i), text = df.columns[i - 1], anchor = 'w')
            
        else :
            tree.column('#{}'.format(i), width = 100, anchor = 'center')
            tree.heading('#{}'.format(i), text = df.columns[i - 1], anchor = 'center')

    for i in range(df.shape[0]) :
        
        tree.insert('', i, values = df.loc[i, :].tolist())
    
    return tree

#########################################################################################
######################################## < action > #####################################

def edit(tree):
   # Get selected item to Edit
   selected_item = tree.selection()[0]
   tree.item(selected_item, text="blub", values=("foo", "bar"))


def delete(tree):
   # Get selected item to Delete
   selected_item = tree.selection()[0]
   tree.delete(selected_item)
