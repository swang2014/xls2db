#=====================================================================================================
#Dojima Solutions
#Stephanie Wang
#=====================================================================================================

import xlrd
import datetime 
import mysql.connector
import sys
import ConfigParser
import os
import platform

#Determines first row of data in the table of the Excel worksheet to avoid reading data from title or heading cells
#Assumes that title rows have consist of mainly empty cells, the heading row proceeds directly after, and the data after that
#Returns the number of the first row of data as an int
def findFirstRow():
    for b in range (0, sheet.nrows):
        k = 0
        for d in range (0, sheet.ncols):
            if sheet.cell(b, d).value is '':
                k = k + 1
        if k < (sheet.ncols - 3):
            firstrow = b + 1
            break
    return firstrow

#Determines the last column of main table of the Excel worksheet in order to avoid reading any extraneous information given in cells outside of the table
#Assumes the table ends at the first empty cell in the heading row
#The number of the last column is returned as an int
def findLastColumn(headingrow):
    for a in range (0, sheet.ncols):
        if sheet.cell(headingrow, a).value == '':
            lastcol = a
            break
        else:
            lastcol = sheet.ncols
    return lastcol

def removeSpace(string):
    for i in range (0, len(string)):
        char = len(string) - (1 + i)
        if string[char] == ' ':
            continue
        else:
            newString = string[:(char + 1)]
            return newString

def findOS():
    system = platform.system()
    #print system
    return system

#Reads the environment variable and creates a path to the Config directory
#The environment variable is a path to the directories containing different configuration documents
def readEnv():
    MYROOT = os.getenv('DOJIMA-XLS2DB-ROOT')
    ConfigDir = MYROOT + '\config'
    print(ConfigDir)
    return ConfigDir  
    
def linuxConfigDir():
    programPath = os.path.realpath(__file__)
    #print programPath
    index = programPath.index('/src/cdsxls2db.py')
    s1 = programPath[:index]
    ConfigDir = s1 + '/config'
    return ConfigDir

#Read command line arguments to find the string, and returns the subsequent argument
def readCommandLine(string):
    totalCount = len(sys.argv)
    index = 0
    while (index < totalCount):
        if sys.argv[index] == string:
            retStr = str(sys.argv[(index + 1)])
            break
        else:
            index += 1
    return retStr

#Takes the given source from the command line argument and reads the correct column map
def chooseFile(ConfigDir, source, system):
    if system== 'Windows': 
        ICEEUColMap = ConfigDir + '\\xls2db-column-mapping\ICE_Europe\ICE_EU.ini'
        ICEUSColMap = ConfigDir + '\\xls2db-column-mapping\ICE_US\ICE-US.ini'
        CMEColMap = ConfigDir + '\\xls2db-column-mapping\CME\CME.ini'
    else:
        ICEEUColMap = ConfigDir + '/xls2db-column-mapping/ICE_EUROPE/ICE_EU.ini'
        ICEUSColMap = ConfigDir + '/xls2db-column-mapping/ICE_US/ICE-US.ini'
        CMEColMap = ConfigDir + '/xls2db-column-mapping/CME/CME.ini'
    if source == "ICE-Europe":
        return ICEEUColMap
    elif source == "ICE-US":
        return ICEUSColMap
    elif source =="CME":
        return CMEColMap
        
#Takes the previously generated column mapping dictionary and a single cell in the heading row of the Excel worksheet to find the corresponding MySQL column
#Returns the name of the corresponding MySQL column
#When no corresponding column is identified, an error message is printed and the returned value is None
def useColMap(colMap, string):
    try:
        newString = removeSpace(string)
        key = newString.lower()
        value = colMap[key]
        #print(value)
        return value
    except:
        print ("According to the dictionary, there is no corresponding column in MySQL for the column:" + string)
        value = None
        return value

#Takes the information from the column map document and returns it into a dictionary        
def readColMap(source, colMapDoc):
    config.read(colMapDoc)
    colMap = dict()
    try:
        for key in config.options(source):
            colMap[key] = config.get(source, key)
        #print("The column map looks like:")
        #print (colMap)
        #print colMap
        return colMap
    except: 
        print("There was some error in putting the INI file data into a Python dictionary")
        return colMap
    
def headingRowList(headingRow, lastcol):
    hlist = []
    for colHeader in range (0, lastcol):
        headingValue = sheet.cell(headingRow,colHeader).value
        newHeadingValue = removeSpace(headingValue)
        hlist.append(newHeadingValue)
    return hlist

#Reads the cells of the heading row in an Excel worksheet and returns all the values in a list
def colHeadingList(hlist, colMap):
    colnames = []
    for i in range (0, len(hlist)):
        value2 = useColMap(colMap, hlist[i])
        #print ("Value 2 is: " + str(value2))
        if value2 == None:
            continue
        else:
            colnames.append(value2)
    colnames.append('DojimaProductType')
    colnames.append('Market')
    return colnames    

def findUnwantedColumns(hlist, coldict):
    badcol = []
    for i in range (0, len(hlist)):
        string = hlist[i] 
        key = string.lower()
        if key in coldict:
            continue
        else:
            badcol.append(i)
    return badcol

#Reads the properties document with the login information, and uses it to connect with MySQL
def connectMySQL(ConfigDir, system):
    if system == 'Windows':
        DBConfigFileName = ConfigDir + "\db\DBConfig.properties"
    else:
        DBConfigFileName = ConfigDir + "/db/DBConfig.properties"
    dbprops = {}
    with open(DBConfigFileName, 'r') as f:
        for line in f:
            line = line.rstrip()

            if "=" not in line: 
                continue
            if line.startswith("#"): 
                continue

            k,v = line.split("=", 1)
            dbprops[k] = v

        #print(dbprops)

        DBUserID = dbprops ['userid']
        DBPwd = dbprops ['password']
        DBHost = dbprops ['host']
        DBdb = dbprops ['database']
        cnx = mysql.connector.connect(user = DBUserID, password = DBPwd, host = DBHost, database = DBdb)
            
        return cnx

#Takes the date given in the Excel worksheet (in float format) and converts it to proper datetime format
#Returns the date 
def floatToDate(excelDate):
    time_tuple = xlrd.xldate_as_tuple(excelDate, 0)
    date = datetime.datetime(*time_tuple)
    #print(date)
    return date

#Takes a string and separates it into multiple strings based on the placement of a single comma and ampersand, and returns the strings as a list
#Used to deal with cells in the coupon rate column, where multiple values are given
def destringify(string):
    if "," in string:          
        L=string.split(", ")
        retL=[]           
        retL.append(L[0])
        L2=L[1].split(" & ")
        for nums in L2:
            retL.append(nums)
        return retL
    elif "&" in string:
        retL=string.split(" & ")
        return retL

#Takes the list of strings generated by the destringify function and converts them to floats. The floats are returned in a list
#This function currently assumes that the coupon rate is given in bps and converts into percentage 
def floatify(L):
    retL=[]
    for num in L:
        retNum=float(num)/100
        retL.append(retNum)
    return retL

#Goes through each cell in the heading row of the Excel worksheet to find the column number of the coupon rate column and returns it as an int
def findCouponColumn(hlist):
    if 'Coupon' in hlist:
        couponcol = hlist.index('Coupon')
    else:
        couponcol = None
    return couponcol
    
#Reads the value from the coupon rate column, deals with the type appropriately, and returns the values as a list of floats
def readCouponRate(couponCol):
    fltrate = []
    #Dealing with coupon rate 
    if couponCol == None:
        fltrate.append(None)   
    elif sheet.cell_type(r,couponCol) == 1:
        strrate = destringify(sheet.cell(r, couponcol).value)
        fltrate = floatify(strrate)
    elif sheet.cell_type(r,couponCol) == 2:
        percoupon = (sheet.cell(r, couponCol).value)/100
        fltrate.append(percoupon)    
    return fltrate

def findClearCol(hlist):
    if "1st Clearing Week" in hlist:
        clearCol = hlist.index("1st Clearing Week")
    else:
        clearCol = None
    return clearCol

#Similar function for Excel worksheet cells in the clearing date column that have multiple dates given per cell
def readClearDate(clearCol):
    cdates = []
    if clearCol == None:
        cdates.append(None)
    elif sheet.cell_type(r,clearCol) == 1:
        stringDates = destringify(sheet.cell(r,clearCol).value)
        for a in range (0, len(stringDates)):
            date_object = datetime.datetime.strptime(stringDates[a], "%d-%b-%y")        
            cdates.append(date_object)
    elif sheet.cell_type(r, clearCol) == 3:
        cdates.append(sheet.cell(r, clearCol).value)
    return cdates

def findProductType(colnames, list):
    if 'Sector' in colnames:
        index =  colnames.index('Sector')
        sectorValue =  removeSpace(list[index])
        if sectorValue == 'Government':
            productType = "Sovereign"
        else:
            productType = "Corporate"       
    else:
        productType = "Index"
    return productType
        
#Takes a string and returns it with double quotes appended around it
#Necessary for SQL syntax in insert statement
def addQuotes(string):
    value = "\"" + string + "\""
    return value
    
#Takes the list of MySQL column names and returns it in a string format for usage in the insert statement
def colListString(strlist):
    stringLength = len(strlist)
    string =str(strlist[0])
    for index in range (1, stringLength):
        stringItem = str(strlist[index])
        string = string + ", " + stringItem
    #print(string)
    return string


def listToString(strlist):
    stringLength = len(strlist)
    if isinstance(strlist[0], str):
        string = "\"" + strlist[0] + "\""
    elif isinstance(strlist[0], unicode):
        string = "\"" + strlist[0] + "\""
    else: 
        string =str(strlist[0])
    for index in range (1, stringLength):
        stringItem = str(strlist[index])
        #print strlist[index]
        #print type(strlist[index])
        if isinstance(strlist[index], str):
            string = string + ", " + "\"" + stringItem + "\""
        elif isinstance(strlist[index], unicode):
            string = string + ", " + "\"" + stringItem + "\""    
        else:
            string = string + ", " + stringItem
            
    #print(string)
    return string

#Open document
config =ConfigParser.ConfigParser()
filename = readCommandLine('-f')
book = xlrd.open_workbook (filename)

source = readCommandLine('-s')

system = findOS()
if system == 'Windows':
    ConfigDir = readEnv()
else:
    ConfigDir = linuxConfigDir()
    
cnx = connectMySQL(ConfigDir, system)

cur = cnx.cursor()

colMapDoc = chooseFile(ConfigDir, source, system)

colMapDict = readColMap(source, colMapDoc)

#Going through all the worksheets
for sheetnumber in range (0, book.nsheets):
    
    sheet = book.sheet_by_index(sheetnumber)
    #print ("We are on sheet #: " + str(sheetnumber))
            
    firstrow = findFirstRow()
    #print("The first row is " + str(firstrow))
    
    lastcol = findLastColumn(firstrow - 1)
    #print ("The last column is " + str(lastcol))
    
    hList = headingRowList((firstrow - 1), lastcol)
    
    couponcol = findCouponColumn(hList)
    #print("The coupon rate column is column: " + str(couponcol))
    
    clearCol = findClearCol(hList)
    #print("The clearing date column is column: " + str(clearCol))
                  
    colnames = colHeadingList(hList, colMapDict)
    #print ("The corresponding column headings in MySQL are: ")
    #print (colnames)  
    
    badcol = findUnwantedColumns(hList, colMapDict)
    #print("The columns that do not match are:")
    #print(badcol)  
    #Reads excel document by making list of cell values by row
    for r in range (firstrow, sheet.nrows):
        i = 0
        fltrate = readCouponRate(couponcol)
        cdates = readClearDate(clearCol)
                   
        for cRates in range (0, len(fltrate)):
            for aDate in range (0, len(cdates)):
                list = []
                #print ("The current coupon rate is: ")
                #print(fltrate[cRates])
            
                for c in range(0, lastcol):       
                    if c in badcol:
                        continue
                    
                    elif sheet.cell(r ,c).value is '':
                        i = i + 1
                        #Dealing with merged cells
                        for j in range (0, sheet.nrows):
                            if sheet.cell((r - j), c).value is not '':
                                list.append(sheet.cell((r - j), c).value)
                                break
                                                                       
                    
                    #Dealing with date types
                    elif sheet.cell_type(r,c) is 3:
                        date = floatToDate(sheet.cell(r,c).value)
                        stringDate = str(date)
                        list.append(stringDate)
                    
                    elif c == clearCol:
                        stringDate = str(cdates[aDate])
                        list.append(stringDate)
                        
                    #Appending coupon rate 
                    elif c == couponcol:
                        list.append(fltrate[cRates])
                        #print("The coupon rate that is being appended is:")
                        #print(fltrate[cRates])
                                    
                    else:
                        list.append(sheet.cell(r,c).value)
                
        
                #Figuring out last row
                if (i+1) >= lastcol:
                    break
        
                else:
                    #print(list)
                    productType = findProductType(colnames, list)
                    list.append(productType)
                    list.append(source)
                    stringColNames = colListString(colnames)
                    stringList = listToString(list)
                    sql = ("INSERT INTO CDS (%s) VALUES (%s)" % (stringColNames, stringList))
                    #print (sql)
                    cur.execute(sql)
    
                    cnx.commit()


#Trying to print out what's in the table
#cur.execute("SELECT * FROM cds")
#rows = cur.fetchall()
#for eachRow in rows:
#    print (eachRow)


cnx.close()



#print ("success?")
