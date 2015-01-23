from sikuli import *
def SaveAsPLGood():
    wait(Pattern("Rebortisbein.png").targetOffset(-4,1),FOREVER)
    waitVanish(Pattern("Rebortisbein.png").targetOffset(-4,1),FOREVER)
    
    #Save 
    #auto run so no need to click generate button
    wait(1)
    click("Export-2.png")
    wait("YZDExcel-2.png",FOREVER)
    click("YZDExcel-2.png")
    wait("SM-2.png",FOREVER)
    click("1352341267969-3.png")   
    wait("IJlyDocument-2.png",FOREVER)
    type(Settings.plname)
    wait(1)
    click(Pattern("1352344598774-1.png").targetOffset(24,-10))
    wait(3)
    if exists(Pattern("1367571573133-2.png").similar(0.90)):
        click(Pattern("yes-2.png").similar(0.90).targetOffset(-29,1))