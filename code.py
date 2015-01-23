from sikuli import *
# works on all platforms
import os
myPath = "Q:\\Henry\\FR\\NewPnL"
Settings.MinSimilarity=0.9

# get the directory containing your running .sikuli
#myPath = os.path.dirname(getBundlePath("dir"))
if not myPath in sys.path: sys.path.append(myPath)

# now you can import every .sikuli in the same directory
App.open('"C:\\Program Files\\Internet Explorer\\iexplore.exe" "https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/reporttree.aspx?subdomain=1&Tab=F"')

#FSO
reportsToRun =  (1,2,3,4,5)
#reportsToRun =  (101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150)
#reportsToRun =  (151,152,153,154,155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,201,202,203,204)
for i in reportsToRun:
    wait (Pattern("StandardReoo.png").similar(0.87),FOREVER)
    mouseMove(Location(416,422))
       
    if i == 1:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45572&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO - China North MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
        
    if i == 2:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45574&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO - China Central MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
		
	if i == 3:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45575&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO China Central YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 4:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45573&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO China North YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 5:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45576&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO China South MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 6:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45577&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO China South YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 7:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45578&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO China MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 8:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45579&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO China YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 9:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45581&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Hong Kong MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 10:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45582&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Hong Kong YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 11:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45584&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Korea MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 12:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45585&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Korea YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 13:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45586&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Singapore MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 14:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45587&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Singapore YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 15:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45590&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Vietnam MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 16:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45592&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Vietnam YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 17:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45593&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Oceania MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 18:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45594&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Oceania YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 19:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45570&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Asia Pacific MTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 20:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45571&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Asia Pacific YTD (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 21:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45676&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Assurance Guangzhou (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 22:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45677&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Assurance Shanghai (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 23:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45678&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Assurance Shenzhen (LC) .xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 24:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45679&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO By Countries (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 25:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45680&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Korea by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 26:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45686&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China All Services  (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 27:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45687&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China Assurance (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 28:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45688&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China Tax (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 29:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45689&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China Advisory (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i ==30:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45690&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China TAS (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 31:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45691&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China Elimination (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 32:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45692&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China Risk Excl. FSRM & AS (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 33:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45693&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China FSRM (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 34:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45694&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China AS (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 35:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=45695&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\FSO Greater China PI (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 36:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47117&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Brunei by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 37:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47118&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Cambodia by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 38:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47119&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Guam by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 39:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47120&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Indonesia by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 40:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47121&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Laos by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 41:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47122&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Malaysia by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 42:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47116&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Maldives by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 43:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47123&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Philippines by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 44:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47124&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Singapore by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 45:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47125&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Sri Lanka by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 46:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47126&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Thailand by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 47:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47127&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Vietnam by SL (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 48:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47130&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ASEAN by Countries (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 49:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47131&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ASEAN by Countries (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 50:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47132&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ASEAN by Countries Exclude FSO & APTC (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 51:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47135&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ASEAN by Countries Exclude FSO & APTC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 52:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47128&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ASEAN by SL (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 53:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47129&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ASEAN by SL Exclude FSO & APTC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 54:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47133&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Sri Lanka by BU (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 55:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47134&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Vietnam by Location (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 56:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47136&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Singapore by SL Exclude FSO & APTC (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 57:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=226&myreportid=47137&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\APAC FSO ITS TP (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 58:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47003&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 59:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47012&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC Australia (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 60:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47007&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 61:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47005&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC China Central (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 62:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47004&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC China North (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 63:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47006&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC China South (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 64:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47008&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC Hong Kong (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 65:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47009&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC Malaysia (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 66:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47010&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - APTC Singapore (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 67:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46958&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Area CBS (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 68:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46952&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - ASEAN (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 69:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46959&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 70:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46953&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Australia (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 71:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46960&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Beijing (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 72:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46981&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Brunei (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 73:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46939&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Brunei (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 74:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46982&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Cambodia (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 75:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46940&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Cambodia (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 76:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46965&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Chengdu (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 77:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46933&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 78:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46971&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - China Central (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 79:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46931&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - China Central (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 80:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46964&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - China North (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 81:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46930&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - China North (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 82:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46975&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - China South (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 83:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46932&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - China South (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 84:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46961&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Dalian (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 85:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46956&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 86
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46997&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
	if i == 87:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46995&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO China Central (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 88:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46994&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO China North (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 89:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46996&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO China South (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 90:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46998&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO Hong Kong (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 92:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47002&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO Oceania (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 91:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46999&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO Korea (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 93:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47000&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO Singapore (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 94:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47001&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO Vietnam (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 95:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46993&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - FSO XC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 96:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46937&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Greater China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 97:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46983&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Guam (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 98:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46941&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Guam (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 99:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46972&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Guangzhou (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 100:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46966&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Hangzhou (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 101:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47015&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Hanoi (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 102:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47014&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Hanoi (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 103:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47017&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - HCMC (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 104:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47016&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - HCMC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 105:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46977&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - HK PRC (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 106:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46976&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Hong Kong (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 107:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46934&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Hong Kong (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 108:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46984&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Indonesia (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 109:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46942&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Indonesia (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 110:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46980&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Korea (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 111:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46938&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Korea (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 112:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46985&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Laos (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 113:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46943&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Laos (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 114:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46986&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Malaysia (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 115:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46944&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Malaysia (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 116:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46987&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Maldives (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 117:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46946&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Maldives (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 118:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46979&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Mongolia (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 119:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46936&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Mongolia (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 120:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46967&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Nanjing (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 121:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46954&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - New Zealand (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 122:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46955&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Oceania (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 123:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46988&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Philippines (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 124:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46947&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Philippines (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 125:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46962&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Qingdao (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 126:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46970&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Shanghai (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 127:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46973&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Shenzhen (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 128:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46989&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Singapore (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 129:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46948&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Singapore (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 130:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46990&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Sri Lanka (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 131:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46949&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Sri Lanka (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 132:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46968&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Suzhou (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 133:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46978&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Taiwan (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 134:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46935&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Taiwan (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 135:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46991&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Thailand (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 136:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46950&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Thailand (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 137:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46963&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Tianjin (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 138:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46992&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Vietnam (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 139:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46951&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Vietnam (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i ==140:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46969&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Wuhan (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 141:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=46974&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by SSL - Xiamen (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 142:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47047&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Additional Svcs Tax by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 143:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47046&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Additional Svcs Tax by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 144:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47916&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\BTA by Markets - APTC (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 145:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47915&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\BTA by Markets - APTC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 146:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47027&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\BTA by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 147:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47026&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\BTA by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 148:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47021&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\BTS by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 149:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47020&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\BTS by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 150:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47075&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\GCR by Markets - APTC (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 151:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47074&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\GCR by Markets - APTC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 152:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47023&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\GCR by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 153:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47022&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\GCR by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 154:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47039&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\HC GM by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 155:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47038&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\HC GM by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 156:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47065&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\HC GM by Markets - Greater China (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 157:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47062&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\HC GM by Markets - Greater China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 158:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47039&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\HC P&R by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 159:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47040&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\HC P&R by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 160:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47041&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\HC P&R by Markets - Greater China (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 161:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47066&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\HC P&R by Markets - Greater China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 162:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47037&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Human Capital by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 163:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47036&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Human Capital by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 164:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47063&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Human Capital by Markets - Greater China (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 165:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47062&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Human Capital by Markets - Greater China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 166:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47043&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Indirect Tax by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 167:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47042&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Indirect Tax by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 168:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47031&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 169:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47030&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 170:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47033&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Core by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 171:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47032&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Core by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 172:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47056&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks Africa by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 173:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47058&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks BeNe by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 174:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47055&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks Brazil by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 175:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47049&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 176:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47057&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks FSO by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 177:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47052&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks Germany by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 178:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47050&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks India by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 179:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47054&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks Japan by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 180:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks UK by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 181:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS Desks US by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 182:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47077&reportTypeID=1" + Key.ENTER)
        wait(5)`
        Settings.plname = "Q:\Henry\FR\Output\ITS TESCM by Markets - APTC (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 183:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47076&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS TESCM by Markets - APTC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 184:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47048&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS TESCM by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 185:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=48676&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS TESCM by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 186:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47035&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS TP by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 187:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47034&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\ITS TP by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 188:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47025&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\PTS by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 189:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47024&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\PTS by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 190:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47029&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\TARAS by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 191:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47028&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\TARAS by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 192:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47072&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - APTC (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 193:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47070&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - ASEAN (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 194:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47069&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - ASEAN (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 195:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47019&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 196:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47018&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 197:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47071&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - FSO (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 198:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47061&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - Greater China (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 199:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47060&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - Greater China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 200:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47078&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Tax by Markets - Malaysia (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 201:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47045&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Transaction Tax by Markets - Asia Pacific (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 202:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47044&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Transaction Tax by Markets - Asia Pacific (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 203:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47061&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Transaction Tax by Markets - Greater China (LC).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
    if i == 204:
        wait(2)
        type("d", KEY_ALT);
        wait(2)
        type("https://mybusiness.ey.net/fr/_layouts/ey.erp.financialreporting.web/ReportViewerPagev1.aspx?reportid=211&myreportid=47060&reportTypeID=1" + Key.ENTER)
        wait(5)
        Settings.plname = "Q:\Henry\FR\Output\Transaction Tax by Markets - Greater China (USD).xls"
        import saveGdPL
        reload(saveGdPL)
        saveGdPL.SaveAsPLGood()
