import c4d
from c4d import gui
from c4d import documents
from c4d import utils
from c4d import Vector, Matrix
import os
import os.path
import xml.etree.ElementTree as etree
import ctypes  # An included library with Python install.
MB_YESNO = 0x04
MB_OK = 0x0
ICON_INFO = 0x40
IDYES = 6
IDOK = 0

Kunde = []
Nummer = []
Blockname = []
BlocknameSort = []
File = []
Shift = []
DirNameC4d = []
DirNameRender = []

def main():
	exportpath = "c:/vbnet"
	user = os.getenv('username') # Windows Loginname
	stlpath= exportpath+"/"+user+"/"
	tree = etree.parse(exportpath+"/"+user+".xml")
	xmlroot = tree.getroot()
	for child in xmlroot:
		Kunde.append(child.find('CUSTOM_PROPERTY_VALUE').text)
		Nummer.append(child.find('DRAWING_NAME').text)
		Blockname.append(child.find('BLOCKNAME').text)
		BlocknameSort.append(child.find('BLOCKNAME').text)
		File.append(child.find('STL_FILENAME').text)
		Shift.append(child.find('SHIFT').text)
		DirNameC4d.append(child.find('SAVEPATHC4DFILE').text)
		DirNameRender.append(child.find('SAVEPATHRENDEROUTPUT').text)
	BlocknameSort.sort()
	BlocknameSort.reverse()
	dirList=os.listdir(stlpath)
	bc = c4d.BaseContainer()
	bc.SetData(c4d.MDATA_UNTRIANGULATE_NGONS, False)
	bc.SetData(c4d.MDATA_UNTRIANGULATE_ANGLE_RAD, .5)

	root = c4d.BaseObject(c4d.Onull) # Create new Null
	root.SetRelPos(c4d.Vector(0))   # Set position
	root.SetName(Kunde[0]+"_"+Nummer[0])
	doc = documents.GetActiveDocument()
	doc.InsertObject(root)
	oldblock = ""
	index = 0
	for item in BlocknameSort:
		if item != oldblock and item != "empty":
			blockroot = c4d.BaseObject(c4d.Onull)
			blockroot.SetRelPos(c4d.Vector(0))
			blockroot.SetName(item)
			doc.InsertObject(blockroot)
			blockroot.InsertUnder(root)
			oldblock = item
	for fname in dirList:
		# print fname
		filepath=stlpath+fname
		c4d.documents.MergeDocument(doc,filepath,c4d.SCENEFILTER_OBJECTS)
		c4d.EventAdd() #refresh the scene.
	for fname in File:
		obj = doc.SearchObject(fname)
		if Blockname[index] == "empty":
			grp = root
		elif Blockname[index] == "__SOLIDS":
			grp = doc.SearchObject(Blockname[index])
		else:
			grp = doc.SearchObject(Blockname[index])
			tmpstring = obj.GetName().replace(grp.GetName() + "_", "")
			tmpstring2 = tmpstring.split("_")
			obj.SetName(tmpstring[len(tmpstring2[0])+1:])
		index = index + 1
		utils.SendModelingCommand(c4d.MCOMMAND_UNTRIANGULATE, list = [obj], mode = c4d.MODIFY_ALL, bc=bc, doc = doc)
		obj.MakeTag(c4d.Tphong)
		obj.SetPhong(True, True, c4d.utils.Rad(20))
		if Shift[0] == "True":
			obj.SetAbsPos(c4d.Vector(-10000,-10000,-10000))
		obj.InsertUnder(grp)
		#tmpstring = obj.GetName().replace(grp.GetName() + "_", "")
		#tmpstring2 = tmpstring.split("_")
		#obj.SetName(tmpstring[len(tmpstring2[0])+1:])
	c4d.EventAdd() #refresh the scene.
	c4d.CallCommand(100004767) # Deselect All
	MessageBox = ctypes.windll.user32.MessageBoxA  # MessageBoxA in Python2, MessageBoxW in Python3
	if(DirNameRender[0] != "" and ("Untitled" in doc.GetDocumentName() or "Ohne Titel" in doc.GetDocumentName())):
		mbid = MessageBox(None, 'In den Rendervoreinstellungen den Dateipfad anpassen...', 'Pfade Anpassen?', MB_YESNO | ICON_INFO)
		if mbid == IDYES:
			rDat = doc.GetActiveRenderData()
			rDat[c4d.RDATA_PATH] = DirNameRender[0]+"\\"+Nummer[0].replace(".", "-")+"-$take"
		else:
			pass
	else:
		print("Keine neues Projekt: "+doc.GetDocumentName())
	num = Nummer[0].replace(".", "-")
	if(os.path.exists(DirNameC4d[0]+"\\"+num)):
		print(DirNameC4d[0]+"\\"+num+" path already exist!!!")
		print("Current Document Name: "+doc.GetDocumentName())
		mbid1 = MessageBox(None, "Ein Projekt mit dieser Nummer existiert bereits: "+num, "Info", MB_OK | ICON_INFO)
		if mbid1 == IDOK:
			print("Datei existiert bereits Meldung bestaetigt.")
			pass
		else:
			pass
	elif("Untitled" in doc.GetDocumentName() or "Ohne Titel" in doc.GetDocumentName()):
		mbid2 = MessageBox(None, "Projekt inkl. Assets Sichern: "+num, "Speichern?", MB_YESNO | ICON_INFO)
		if mbid2 == IDYES:
			print(DirNameC4d[0]+"\\"+num)
			missingAssets = []
			assets = []
			c4d.documents.SaveProject(doc, c4d.SAVEPROJECT_ASSETS | c4d.SAVEPROJECT_SCENEFILE | c4d.SAVEPROJECT_DIALOGSALLOWED, DirNameC4d[0]+"\\"+num, assets, missingAssets)
		else:
			pass
	else:
		print("Projekt ist bereits an einem anderen Ort gespeichert. Current Document Name: "+doc.GetDocumentName())
if __name__=='__main__':
	main()