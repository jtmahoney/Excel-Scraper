import easygui as g
import csv
import openpyxl as xl
import os
import fnmatch
import sys

data=[]
xlsfiles=[]
fieldNames = ['Directory', 'file type "quote"', 'Sheet Number']
raw_inputs = []
walk_dir = ''
file_quote = ''
sheet_num = ''
cell_loc = []


def get_num_inputs(x):
    msg = "Enter the number of cells (per excel file) you wish to copy information from"
    title = "Number of Inputs"
    numFieldName = ""
    fieldValue = []
    fieldValue = g.integerbox(msg, title, numFieldName)
    print (fieldValue)

    for i in range(0,fieldValue):
        v = i+1
        fieldNames.append('Cell %s:' %v)
        
    

    
def get_inputs(x):
    msg = "Enter the directory path and cell numbers you would like to export"
    title = "Extract from Directory"
    fieldValues = []  # we start with blanks for the values
    fieldValues = g.multenterbox(msg, title, fieldNames)

#    make sure that none of the fields were left blank
    while 1:
        if fieldValues == None: break
        errmsg = ""
        for i in range(len(fieldNames)):
            if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
        if errmsg == "": break # no problems found
        fieldValues = multenterbox(errmsg, title, fieldNames, fieldValues)
    print ("Reply was:", fieldValues)

    for i in fieldValues:
        raw_inputs.append(i)
    

def make_var(x):
    global walk_dir
    walk_dir = x[0]
    global file_quote
    file_quote = x[1]
    global sheet_num
    sheet_num = x[2]
    for r in x[3:]:
        cell_loc.append(r)


def find_files(x): #Populates files list with file names(location)of excel files
    path = ('*%s*' %file_quote) #filter for excel files.  Add info for additioanly filtering
    print (path)
    print (x)
    all_files = []
    print (all_files)
    for subdir, dirs, files in os.walk(x):
        for file in files:
            all_files.append(os.path.join(subdir,file))
    
    y = fnmatch.filter(all_files, path)
    for i in y:
        xlsfiles.append(i)
    
     

def getx(x):  #Opens workbook and copies cell specific data to data list
    for z in x:
        try:
            rb1 = xl.load_workbook(str(z))
            sh = rb1[int(sheet_num)-1]
            NewList = []
            for cells in cell_loc:#specifies which cells to copy
                #if (sh.cell_value(rowx=(r-1), colx=(c-1))) != None:
                try:
                    NewList.append(sh.cell(str(cells)))
                except:
                    NewList.append("")
               
            data.append(NewList)
        except:
            print ('~~Error with file:' + z)
    
   
def wcsv(dat): #This function takes the data list and outputs to a csv file
  
    ofile = open('Target.csv', 'a')
    writer = csv.writer(ofile, delimiter=',', quoting=csv.QUOTE_NONE)

    for i in dat:
        print (i)
        csvdat = i
        try:
            writer.writerow(csvdat)
        except:
            writer.writerow("")
    #ofile.close()

       
get_num_inputs(0)    
get_inputs(0)
print (raw_inputs)
make_var(raw_inputs)
print (walk_dir)
print (file_quote)
print (cell_loc)
print('walk_dir = ' + walk_dir)
find_files(walk_dir)
print (xlsfiles)
getx(xlsfiles)
print (data)
wcsv(data)

sys.exit()
