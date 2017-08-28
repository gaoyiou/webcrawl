import arcpy, os

'''
#export one mxd to pdf
inputmxd = r"C:\Users\agao\Desktop\EWT\Map17-0120 - Crossings Reference Map V2.mxd"
outputfolder = r"C:\Users\agao\Desktop\export\\"

mxd = arcpy.mapping.MapDocument(inputmxd)

noextension = inputmxd[0:-4]
basename = os.path.basename(noextension)
df = arcpy.mapping.ListDataFrames(mxd)

layer = arcpy.mapping.ListLayers(mxd, "Crossings")[0]
if layer.supports("LABELCLASSES"):
    print "supported"
    for lblclass in layer.labelClasses:
        if lblclass.className == "Greenwich Wind Access Road" or lblclass.className == "Greenwich Wind Collection Line":
            lblclass.showClassLabels = True
        elif lblclass.className == "Greenwich Wind Collection Line":
            lblclass.showClassLabels = True
        else:
            lblclass.showClassLabels = False
layer.showLabels = True

addLayer = arcpy.mapping.Layer(r"C:\Users\agao\Desktop\export\EWT_NE_TLine_AdditionalEasement_Scratch_CA - 20160713.lyr")

#update a layer's symbology using a layer file
for eachdf in df:
    arcpy.mapping.AddLayer(eachdf, addLayer, "TOP")
    
    updateLayer = arcpy.mapping.ListLayers(mxd, "Temporary_AR_Area_20170411", eachdf)
    if len(updateLayer) > 0:
        updateLayer = updateLayer[0]
    else:
        continue
    sourceLayer = arcpy.mapping.Layer(r"C:\Users\agao\Desktop\export\Temporary_AR_Area_20170411.lyr")
    arcpy.ApplySymbologyFromLayer_management (updateLayer, sourceLayer)

    #arcpy.mapping.UpdateLayer(eachdf, updateLayer, sourceLayer, True)

#change text in a map
for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
    elm.text = elm.text.replace('Draft, Preliminary Design Not For Construction','Preliminary Design For Construction')
   


#Move layout element
for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
    elm.text = elm.text.replace('File:','Time:')
    if elm.text == '<dyn type="document"property="name"/>':
        elm.delete()
    elm.text = elm.text.replace('<dyn type="date" format="yyy-MM-dd"/>','<dyn type="date" format="yyy-MM-dd"/>\n<dyn type="time" format=""/>')




for elm in arcpy.mapping.ListLayoutElements(mxd, ""):
    if round(elm.elementPositionX,1) == 19.2 and round(elm.elementPositionY,1) == 8.1:
        elm.elementPositionX = 1
        elm.elementPositionY = 9
       

#export map to PDF
#mxd.saveACopy(r"C:\Users\agao\Desktop\savefolder\\Map16-0137 - EWT6001_TARV3.mxd")
arcpy.mapping.ExportToPDF(mxd, outputfolder + basename + ".pdf")
print 'Finish exporting'


#export multiple mxds to pdf
mxdPath = r"C:\Users\agao\Desktop\EWT\\"
newmxdfolder = r"C:\Users\agao\Desktop\savefolder\\"
outputfolder = r"C:\Users\agao\Desktop\export\\"

for fileName in os.listdir(mxdPath):
    mxd = arcpy.mapping.MapDocument(mxdPath + fileName)
    basename = fileName[0:-4]
    df = arcpy.mapping.ListDataFrames(mxd)
    #update a layer's symbology using a layer file
    for eachdf in df:
        updateLayer = arcpy.mapping.ListLayers(mxd, "Temporary_AR_Area_20170411", eachdf)
        if len(updateLayer) > 0:
            updateLayer = updateLayer[0]
        else:
            continue
        sourceLayer = arcpy.mapping.Layer(r"C:\Users\agao\Desktop\export\Temporary_AR_Area_20170411.lyr")
        arcpy.ApplySymbologyFromLayer_management (updateLayer, sourceLayer)
        #arcpy.mapping.UpdateLayer(eachdf, updateLayer, sourceLayer, True)

    #change text in a map
    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
        elm.text = elm.text.replace('Parcel Data from','Crown Data from')

    #save to a new mxd and export to PDF
    mxd.saveACopy(os.path.join (newmxdfolder + basename + ".mxd"))
    arcpy.mapping.ExportToPDF(mxd, outputfolder + basename + ".pdf")
    print basename

print 'All MXDs saved'
print 'Finish exporting'


#Set file name and remove if it already exists
pdfPath = r"C:\Users\agao\Desktop\export\test123.pdf"
if os.path.exists(pdfPath):
    os.remove(pdfPath)
'''

pdfFolder = r"C:\Users\agao\Desktop\export\New folder\\"
for pdffile in os.listdir(pdfFolder):
    pdf = arcpy.mapping.MapDocument(pdfFolder + pdffile)


pdfPath = r"C:\Users\agao\Desktop\export\test123.pdf"

pdfDoc = arcpy.mapping.PDFDocumentOpen(pdfPath)
#pdfDoc.deletePages(1)
pdfDoc.insertPages(r"C:\Users\agao\Desktop\export\Map17-0120 - Crossings Reference Map V2.pdf", 2)

'''#append pages
pdfDoc.appendPages(r"C:\Users\agao\Desktop\export\Map16-0137 - EWT8009_TARV3.pdf")
'''
#Commit changes and delete variable reference
pdfDoc.saveAndClose()
del pdfDoc
print "Done"
