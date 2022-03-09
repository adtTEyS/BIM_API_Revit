# -*- coding: utf-8 -*-

#pyRevit info
__title__ = """Sincronización 
habitaciones Revit-Excel"""
__doc__ = ''
__author__  = 'ADT'

#_________________iMPORTACIONES_______________

#Importar RevitAPI
import clr
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import*
from Autodesk.Revit.UI import *

#Importar forms
from pyrevit import forms

#Importar DocumentManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
import sys
import System
import time 
from System import Array
from System.Collections.Generic import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#_________________iNTERACCIÓN 1_SELECCIÓN DE HABITACIONES_______________

Opciones = ['Selección', 'Todas las habitaciones']
SelRooms = []
SelRooms = forms.SelectFromList.show(Opciones, button_name='Selección de habitaciones')

if SelRooms == 'Selección':
	selection = uidoc.Selection
	selection_ids = selection.GetElementIds()
 
	rooms = []
 
	for element_id in selection_ids:
			rooms.append(doc.GetElement(element_id))

elif SelRooms == 'Todas las habitaciones':
	rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()

#___________________VINCULACIÓN CON BASE DE DATOS EN EXCEL_______________

t = Transaction(doc, 'Read Excel spreadsheet.')
t.Start()

#Accessing the Excel applications.
xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')

#Worksheet, Row, and Column parameters
worksheet = 1
rowStart = 4
rowEnd = 200
columnStart = 1
columnEnd = 5
data = []

#Extracción de datos del excel
EData=[]
data3=[]
EData0=[]

for i in range(rowStart, rowEnd):
	data=[]
	data1=[]
 
	for j in range(columnStart, columnEnd):
		data = xlApp.Worksheets(worksheet).Cells(i, j).Text

		data1.append(data)
	EData0.append(data1)
 
#Limpieza de los datos extraidos del excel
for E in EData0:
    
    if E[0] != '':
        EData.append(E)

t.Commit()

#Busqueda y seleccion de parametro para leer y comparar con la base de datos
room = rooms [0]
p = room.GetOrderedParameters()
parametros= [a for a in p]

nombre_parametros = []

for p in parametros:
    nombre_parametros.append(p.Definition.Name)

#PARÁMETRO 1
items = nombre_parametros
Nombre_parametro_a_consultar = forms.SelectFromList.show(items, button_name='Selecciona el parámetro de la habitación que contiene el estilo de habitación')

#___________________RELLENAR PARÁMETROS DE REVIT CON LA BASE DE DATOS DE EXCEL_______________
t = Transaction(doc,"Creación de suelos")
t.Start()

count=0
count1=0

roomsId=set()
roomsInId=set()
roomsOutId=set()
lrooms=[]
Output=[]
roomsMod=set()

for r, Ed in [(r,Ed) for r in rooms for Ed in EData]:
    
    if Ed[0]==(rooms[count].LookupParameter(Nombre_parametro_a_consultar).AsString()):
        roomsInId.add(rooms[count].Id.ToString())
        roomsMod.add(rooms[count].Id.ToString())
        
        #Param Acabado de muro
        WallParam = rooms[count].LookupParameter('Acabado de muro')
        WallParam.Set(Ed[1])

        #Param Acabado de suelo
        WallParam = rooms[count].LookupParameter('Acabado del suelo')
        WallParam.Set(Ed[2])

        #Param Acabado de techo
        WallParam = rooms[count].LookupParameter('Acabado del techo')
        WallParam.Set(Ed[3])

    else:
        roomsId.add(rooms[count].Id.ToString())
    
    count1 +=1

    if len(EData)==count1:
        count +=1
        count1=0

    roomsOutId = [x for x in roomsId if x not in roomsInId]
    lroomsIn=len(roomsMod)
    lrooms=len(roomsOutId)

t.Commit() 

# output a model's information
message = """Proceso finalizado.

Habitaciones modificadas escribiendo los parametros: {lroomsIn}

Habitaciones no modificadas: {lrooms}

Las habitaciones: {roomsOutId} ; no tienen actualizado el parametro de habitación departamento y por tanto no han podido actualizarse los codigos de acabado.""".format(lrooms=lrooms,roomsOutId=roomsOutId,lroomsIn=lroomsIn)

print(message)