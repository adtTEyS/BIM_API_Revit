# -*- coding: utf-8 -*-

#pyRevit info
__title__ = """Comprobación 
de habitaciones"""
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
placed, unplaced = [], []

for r in rooms:
	if r.Area > 0:
		placed.append(r.Id.ToString())

	else:
		unplaced.append(r.Id.ToString())

tuplaplaced=(placed)

tuplaunplaced=(unplaced)

errores=len(unplaced)

# output a model's information
message = """
LIST OF ROOMS


	Habitaciones Correctamente Colocadas:   {tuplaplaced}

	Habitaciones Mal Colocadas:             {tuplaunplaced}

	{errores} ERRORES""".format(tuplaplaced=tuplaplaced,tuplaunplaced=tuplaunplaced,errores=errores)

print(message)
