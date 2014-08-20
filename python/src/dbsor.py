import os
import mysql.connector
import platform



def linuxConfigDir():
    programPath = os.path.realpath(__file__)
    #print programPath
    index = programPath.index('/src/cdsxls2db.py')
    s1 = programPath[:index]
    ConfigDir = s1 + '/config'
    return ConfigDir

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
    

ConfigDir = linuxConfigDir()

cnx = connectMySQL(ConfigDir, system)
cur1 = cnx.cursor()
cur2 = cnx.cursor()
cur3 = cnx.cursor()

sql1 = "INSERT INTO IRS (SELECT IRS.DojimaProductType, IRS.SubProductType, 'sor', IRS.ClearingHouse, IRS.DojimaSymbol, IRS.DojimaSymbol, IRS.Currency, IRS.MaturityDate, IRS.Description, NULL, NULL, IRS.IsElectronicallyTradable, IRS.IsActive, NULL, NULL, NULL, NULL, NULL, NULL, NULL, IRS.Tenor, NULL, IRS.SwapLeg1Type, IRS.SwapLeg1Frequency, IRS.SwapLeg1DayCount, IRS.SwapLeg1Rate, IRS.SwapLeg2Type, IRS.SwapLeg2Frequency, IRS.SwapLeg2DayCount, IRS.SwapLeg2Rate, NULL, NULL FROM (SELECT DojimaProductType, DojimaSymbol, count(*), MIN(Description) as DescriptionFlag, sum(CASE market when 'sor' then 1 else 0 end) as sor FROM IRS GROUP BY DojimaProductType, DojimaSymbol HAVING count(*) > 1 AND sor = 0) A inner join IRS ON A.DojimaProductType = IRS.DojimaProductType AND A.DojimaSymbol = IRS.DojimaSymbol AND A.DescriptionFlag = IRS.Description);"
cur1.execute(sql1)

sql2 = "INSERT INTO Bonds (SELECT Bonds.DojimaProductType, Bonds.SubProductType, 'sor', Bonds.ClearingHouse, Bonds.DojimaSymbol, Bonds.DojimaSymbol, Bonds.Currency, Bonds.MaturityDate, Bonds.Description, NULL, NULL, Bonds.IsElectronicallyTradable, Bonds.IsActive, NULL, NULL, NULL, NULL, NULL, NULL, NULL, Bonds.Tenor, Bonds.CouponRate, NULL, NULL, NULL, NULL, NULL FROM (SELECT DojimaProductType, DojimaSymbol, count(*), MIN(Description) as DescriptionFlag, sum(CASE market when 'sor' then 1 else 0 end) as sor FROM Bonds GROUP BY DojimaProductType, DojimaSymbol HAVING count(*) > 1 AND sor = 0) A inner join Bonds ON A.DojimaProductType = Bonds.DojimaProductType AND A.DojimaSymbol = Bonds.DojimaSymbol AND A.DescriptionFlag = Bonds.Description);"
cur2.execute(sql2)

sql3 = "INSERT INTO Combos (SELECT Combos.DojimaProductType, Combos.SubProductType, 'sor', Combos.ClearingHouse, Combos.DojimaSymbol, Combos.DojimaSymbol, Combos.Currency, Combos.MaturityDate, Combos.Description, NULL, NULL, Combos.IsElectronicallyTradable, Combos.IsActive, NULL, NULL, NULL, NULL, NULL, NULL, NULL, Combos.NumberOfLegs, Combos.Leg1AssetClass, Combos.Leg2AssetClass, Combos.Leg3AssetClass, Combos.Leg1Market, Combos.Leg2Market, Combos.Leg3Market, Combos.Leg1DojimaSymbol, Combos.Leg2DojimaSymbol, Combos.Leg3DojimaSymbol, Combos.Leg1Side, Combos.Leg2Side, Combos.Leg3Side, Combos.Leg1Ratio, Combos.Leg2Ratio, Combos.Leg3Ratio, NULL, NULL FROM (SELECT DojimaProductType, DojimaSymbol, count(*), MIN(Description) as DescriptionFlag, sum(CASE market when 'sor' then 1 else 0 end) as sor FROM Combos GROUP BY DojimaProductType, DojimaSymbol HAVING count(*) > 1 AND sor = 0) A inner join Combos ON A.DojimaProductType = Combos.DojimaProductType AND A.DojimaSymbol = Combos.DojimaSymbol AND A.DescriptionFlag = Combos.Description);"
cur3.execute(sql3)

cnx.commit() 
cnx.close()
