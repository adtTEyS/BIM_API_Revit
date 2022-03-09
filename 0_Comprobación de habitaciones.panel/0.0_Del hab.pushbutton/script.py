# -*- coding: utf-8 -*-

#pyRevit info
__title__ = """Eliminación de
habitaciones erroneas"""
__doc__ = ''
__author__  = 'ADT'

import clr
#Importar RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import*
#Importar DocumentManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Collecting all the rooms
collector   = FilteredElementCollector(doc)
rooms = collector.OfCategory(BuiltInCategory.OST_Rooms).ToElements()

# Filter rooms to those with area more than 0

i=[]
Id=[]

t=Transaction(doc, "Borrado de habitaciones mal colocadas")
t.Start()

count=0

for r in rooms:
	if r.Area > 0:
		pass
	else:
		Id.append(r.Id.ToString())
		doc.Delete(r.Id)
		count+=1
		i.append(count)
		
i=len(i)
t.Commit()

# output a model's information
message = """	{i} Habitaciones erroneas a eliminar
	Habitaciones eliminadas: {Id} """.format(i=i,Id=Id)

print(message)



	
