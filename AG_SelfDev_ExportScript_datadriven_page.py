import arcpy, os, string

inputmxd = r"C:\Users\agao\Desktop\EWT\Map17-0127.mxd"
outputfolder = r"C:\Users\agao\Desktop\export\data driven page export\\"

mxd = arcpy.mapping.MapDocument(inputmxd)

noextension = inputmxd[0:-4]
basename = os.path.basename(noextension)

#Cycle through data driven pages and export as a new MXD
#for pageNum in range(1, mxd.dataDrivenPages.pageCount + 1):
for pageNum in range(20, 35):
    mxd.dataDrivenPages.currentPageID = pageNum
    pageName = mxd.dataDrivenPages.pageRow.Client_ID
    df = arcpy.mapping.ListDataFrames(mxd,"Layers")[0]

    #Check for proper scale rouding
    
    if df.scale < 1000:
        df.scale = round(df.scale,-2)
    elif df.scale > 1000:
        df.scale = round(df.scale/500) * 500


#export map to new mxd and PDF
    mxd.saveACopy(outputfolder + basename + " - " + str(pageName)+".mxd")
    arcpy.mapping.ExportToPDF(mxd, outputfolder + basename + " - " + str(pageName)+".pdf")
    print pageName


print 'Finish exporting'


