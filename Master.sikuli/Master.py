from sikuli import *
# works on all platforms
import os
myPath = "Q:\Donald\FR\FR ADV\NewPNL"
Settings.MinSimilarity=0.9
Settings.MonthVal = " FY14 Feb"   #<--------------------Don't forget to change this#########################################
Settings.LocationCBS = "Q:\Donald\FR\FR ADV\Output\\"

# get the directory containing your running .sikuli
#myPath = os.path.dirname(getBundlePath("dir"))
if not myPath in sys.path: sys.path.append(myPath)

# now you can import every .sikuli in the same directory
App.open('"C:\\Program Files\\Internet Explorer\\iexplore.exe" "https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/reporttree.aspx?subdomain=1&Tab=F"')

###########################################ADV PnL Reports#####################################################
reportsToRun = range(83,95)
#reportsToRun in range(48)
###between 10 and 47 is y in range(10,48)
for i in reportsToRun:
    wait (Pattern("StandardReoo.png").similar(0.87),FOREVER)
    mouseMove(Location(416,422))
    if i == 1:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=95965&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC BRS Core by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 2:
		wait(2)
		type("d", KEY_ALT);
		wait(2)
		type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=84940&reportTypeID=1" + Key.ENTER)
		wait(5)
		Settings.plname = Settings.LocationCBS +"ADV APAC by Markets"+Settings.MonthVal + ".xls"
		import saveGdPL
		reload(saveGdPL)
		saveGdPL.SaveAsPLGood()	
    if i == 3:
		wait(2)
		type("d", KEY_ALT);
		wait(2)
		type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=85885&reportTypeID=1" + Key.ENTER)
		wait(5)
		Settings.plname = Settings.LocationCBS +"ADV APAC incl FSO by SSL"+Settings.MonthVal + ".xls"
		import saveGdPL
		reload(saveGdPL)
		saveGdPL.SaveAsPLGood()
    if i == 4:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=69640&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC ITRA"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 5:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=80602&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC PI FSO"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 6:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=68575&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC PI incl FSO"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 7:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=86185&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC PI"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 8:
		wait(2)
		type("d", KEY_ALT);
		wait(2)
		type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=96462&reportTypeID=1" + Key.ENTER)
		wait(5)
		Settings.plname = Settings.LocationCBS +"ADV APAC Risk Core by SSLs"+Settings.MonthVal + ".xls"
		import saveGdPL
		reload(saveGdPL)
		saveGdPL.SaveAsPLGood()
    if i == 9:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=71483&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC Risk FSO by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 10:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=91504&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC Risk FSO by SSLs"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 11:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=65746&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC Risk incl FSO by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 12:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=71483&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV APAC Risk incl FSO by SSLs"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 13:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=95968&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN Core by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 14:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=96167&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN Core by SSLs"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 15:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=95969&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN ITRA by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 16:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95972&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN ITRA by Markets_until GM before PNC"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 17:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=71632&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN PI by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 18:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95975&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN PI by Markets_until GM before PNC"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 19:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=71632&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN PI exFSO"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 20:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=95970&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN Risk by Markets"+ Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 21:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95974&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ASEAN Risk by Markets_until GM before PNC"+ Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 22:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=73112&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV FSO by Markets"+ Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 23:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=70494&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV FSO by SSLs"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 24:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=71365&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GC Core by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 25:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=71365&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GC exFSO by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 26:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=91121&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GC FSO AS by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 27:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=73743&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GC FSO by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 28:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=71141&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GC FSO by SSL "+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 29:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=91120&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GC FSO FSRM by Markets  "+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i ==30:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=91122&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GC FSO Risk (excl FSRM & AS) by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 31:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=74104&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GC incl FSO by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 32:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=96166&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV GU by SSLs_until GM before PNC"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 33:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=66967&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV HK Core By SSLs"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 34:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95978&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV IN by SSLs_until GM before PNC"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 35:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=79851&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV ITRA GC Core by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 36:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95983&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV LK by SSLs_until GM before PNC"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 37:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95982&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV MV by SSLs_until GM before PNC"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 38:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95977&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV MY by SSLs_until GM before PNC"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 39:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=95961&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV Oceania Core by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 40:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=94376&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV Other Advisory GC ex FSO by Markets"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 41:	
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=94375&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV Other Advisory GC FSO by Markets "+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 42:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95979&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV PH by SSLs_until GM before PNC"+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 43:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=90962&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV PI Commercial GC Core by Markets  "+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 44:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=68828&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV PI Energy GC Core by Markets  "+Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 45:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=65353&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV PI GC Core by Markets " + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 46:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=76300&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV PI GC FSO by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 47:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=79028&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV PI GC incl FSO by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 48:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=90961&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS +"ADV PI General GC Core by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 49:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=74648&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "ADV PI Public GC Core by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 50:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=80554&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "ADV PI TCE GC Core by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 51:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=70732&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "ADV Risk GC Core by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 52:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95976&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "ADV SG by SSLs until GM before PNC" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 53:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95980&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "ADV TH by SSLs until GM before PNC" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)	
    if i == 54:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=401&myreportid=95981&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "ADV VN by SSLs until GM before PNC" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 55:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=89739&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "AS FSO APAC by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 56:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=89447&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Driver ADV - APAC Risk (incl FSO) " + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 57:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95964&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Driver ADV - APAC Risk AS" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 58:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=89454&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Driver ADV - APAC ITRA" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 59:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95963&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Driver ADV - APAC Risk BRS" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 60:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=96168&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Driver ADV - APAC Risk FSRM" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 61:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95925&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO - AS" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 62:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95921&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO  - BRS" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 63:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95927&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO  - FSO Korea" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 64:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95929&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO  - FSO Oceania" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 65:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95930&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO  - FSO Singapore" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 66:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95928&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO  - FSO Vietnam" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 67:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95923&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO  - FSRM" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 68:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95920&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO  - ITRA" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 69:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95926&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers ADV FSO  - PI" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 70:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95919&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers - ADV FSO" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 71:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=94367&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers - ADV GC Core" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 72:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=94368&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers - ADV GC FSO" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 73:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=94366&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers - ADV GC incl FSO" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 74:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95931&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers - ADV Korea Core " + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 75:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=95962&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers - APAC ADV" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 76:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=92500&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers  - ASEAN ITRA" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 77:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=94889&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers  - ASEAN PI" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 78:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=92499&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers  - ASEAN Risk" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 79:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=373&myreportid=92502&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Finance Business Drivers  - ASEAN" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 80:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=89451&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "FSRM FSO APAC by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 81:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=399&myreportid=73112&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "Risk FSO APAC by Markets" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 82:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=70837&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report GEO All" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 83:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=76321&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report GEO ASU" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 84:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=65901&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report GEO TAX" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 85:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=71385&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report GEO ADV" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 86:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=77578&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report GEO TAS" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 87:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=66042&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report GEO CBS" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 88:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=94377&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report GEO ACR" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 89:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=77068&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report Mngm ALL" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i ==90:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=66639&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report Mngm ASU" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 91:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=72227&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report Mngm TAX" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 92:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=77476&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report Mngm ADV" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 93:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=66714&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report Mngm TAS" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 94:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=74723&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report Mngm CBS" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
    if i == 95:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=413&myreportid=94378&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = Settings.LocationCBS + "MTD YTD APAC Report Mngm ACR" + Settings.MonthVal + ".xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        wait(2)
		
	####8 9 19 22 48 49 50 51
    if exists("mOinFolder-10.png"):
        print("1")
        waitVanish("mOinFolder-11.png",FOREVER)
        print("2")
    print("3")
    wait(Pattern("ReportsFinan-8.png").targetOffset(-3,1),FOREVER)
    print("4")
    click(Pattern("ReportsFinan-9.png").targetOffset(0,1))    