# -*- coding: utf-8 -*-

#pyRevit info
__title__ = """Sincronización tipos 
de acabado Revit-Excel"""
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

import itertools

#_________________UTILIDADES_______________

#Conversión de unidades
ft = 304.8 # 1ft = 304.8mm

#_________________iNTERACCIÓN 1_SELECCIÓN DE HABITACIONES_______________

Opciones = ['Selección', 'Todas las habitaciones']
SelRooms = []
SelRooms = forms.SelectFromList.show(Opciones, button_name='Selección de habitaciones')
rooms = []

if SelRooms == 'Selección':
	selection = uidoc.Selection
	selection_ids = selection.GetElementIds()
    
	for element_id in selection_ids:
			rooms.append(doc.GetElement(element_id))

elif SelRooms == 'Todas las habitaciones':
	rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()

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

#_________________iNTERACCIÓN 2_SELECCIÓN DE ACABADO A MODELAR_______________

Opciones = ['Acabado de techo', 'Acabado de suelo', 'Acabado de muro']
Acabados = []
Acabados = forms.SelectFromList.show(Opciones, button_name='Selección de acabados')

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

#_________________iNTERACCIÓN 3_SELECCIÓN DE TIPO DE ACABADO_______________

Tipo= []
TId= []
TipoType = []
tipos = []
roomsId=set()
roomsInId=set()
roomsOutId=set()
lrooms=[]
TEd = set()
FEd = set()
AllEd = set()
AllP= set()
TP = set()
FP = set()

#_________________PREPARACIÓN DE ACABADOS DE TECHO
if Acabados == 'Acabado de techo':
    types = FilteredElementCollector(doc).OfCategory(
                                                BuiltInCategory.OST_Ceilings
                                                ).WhereElementIsElementType()
    
    #Para parámetros de techos:
    ceilings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()
    ceiling = ceilings[0]

    #PARÁMETRO 2
    with forms.WarningBar(title='Selecciona el parámetro de CONSULTA donde está el código de acabado de techo'):
        selected_parameters = forms.select_parameters(ceiling)[0]
        Nombre_parametro_consulta = selected_parameters.name
        
    Lrooms=[]
    Idroom=[]
    levelrooms1=[]
    rooms1=[] 
    EData1=[]
    rooms2=[]
    Lrooms2=[]
    
    
    for r, Ed in [(r,Ed) for r in rooms for Ed in EData]:
        if Ed[0]== r.LookupParameter(Nombre_parametro_a_consultar).AsString():
            Lrooms.append(r.Level)
            rooms1.append(r)
            EData1.append(Ed)
            
    types1=[]
    TId=[]
    
    count=0
    
    for Ed in EData1:
        ATS = Ed[3].split("/")
        ATS0 = ATS[0]
        
        AllEd.add(Ed[3])
        AllP.add(Ed[0])
        
        for t in types:
            roomsId.add(rooms1[count].Id.ToString())
            
            if ATS0 == t.LookupParameter(Nombre_parametro_consulta).AsString():
                TId.append(t.Id)
                types1.append(t)
                rooms2.append(rooms1[count])
                Lrooms2.append(rooms1[count].Level.Id)
                roomsInId.add(rooms1[count].Id.ToString())
                TEd.add(Ed[3])
                TP.add(Ed[0])
                            
        count+=1

#_________________PREPARACIÓN DE ACABADOS DE SUELO        
elif Acabados == 'Acabado de suelo':
    types = FilteredElementCollector(doc).OfCategory(
                                                BuiltInCategory.OST_Floors
                                                ).WhereElementIsElementType()
    
    #Para parámetros de suelos:
    floors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
    floor = floors[0]

    #PARÁMETRO 2
    with forms.WarningBar(title='Selecciona el parámetro de CONSULTA donde está el código de acabado en el suelo'):
        selected_parameters = forms.select_parameters(floor)[0]
        Nombre_parametro_consulta = selected_parameters.name
        
    Lrooms=[]
    Idroom=[]
    levelrooms1=[]
    rooms1=[] 
    EData1=[]
    
    for r, Ed in [(r,Ed) for r in rooms for Ed in EData]:

        if Ed[0] == r.LookupParameter(Nombre_parametro_a_consultar).AsString():
            Lrooms.append(r.Level)
            rooms1.append(r)
            EData1.append(Ed)
                
    types1=[]        
    TId=[]
    Lrooms2=[]
    rooms2=[]
    
    count=0
    count1=0
    
    for Ed in EData1:
        ASS = Ed[2].split("/")      
        ASS0 = ASS[0]   
        
        for t in types:
            
            if ASS0 == t.LookupParameter(Nombre_parametro_consulta).AsString():
                
                TId.append(t.Id)
                types1.append(t)
                rooms2.append(rooms1[count])
                Lrooms2.append(rooms1[count].Level.Id)
                
        count+=1

#_________________PREPARACIÓN DE ACABADOS DE MURO            
elif Acabados == 'Acabado de muro':
    types = FilteredElementCollector(doc).OfCategory(
                                                BuiltInCategory.OST_Walls
                                                ).WhereElementIsElementType()

    #Para parámetros de muros:
    Walls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
    Wall = Walls[0]

    #PARÁMETRO 2
    with forms.WarningBar(title='Selecciona el parámetro de CONSULTA donde está el código de acabado en el suelo'):
        selected_parameters = forms.select_parameters(Wall)[0]
        Nombre_parametro_consulta = selected_parameters.name
        
    Lrooms=[]
    Idroom=[]
    levelrooms1=[]
    rooms1=[]
    rooms2=[]
    EData1=[]
    ValueWidth=[]

    #Buscar las coincidencias entre la base de datos de excel con el estilo de habitacion en revit
    for r, Ed in [(r,Ed) for r in rooms for Ed in EData]:
        if Ed[0]== r.LookupParameter(Nombre_parametro_a_consultar).AsString():
            rooms1.append(r) #Habitaciones existentes como estilo de habitacion
            EData1.append(Ed) #Base de datos de excel filtrada
            
    #Buscar las coincidencias entre la base de datos de excel filtrada en el paso anterior y los tipos de acabados creados en revit
    types1=[]

    count=0
    for Ed in EData1:
        AWS = Ed[1].split("/")      
        AWS0 = AWS[0] 
        
        for t in types:
            if AWS0 == t.LookupParameter(Nombre_parametro_consulta).AsString():

                rooms2.append(rooms1[count]) #Habitaciones filtradas a la par con los tipos
                Lrooms.append(rooms1[count].Level) #Niveles de habitación existentes como estilo de habitación

                TId.append(t.Id) #Id del tipo de muro coincidente con la BD de excel y con el tipo de revit
                types1.append(t) #Tipo de acabado de muro coincidente con la BD de excel y con el tipo de revit
                ValueWidth.append(t.Width) #Ancho de acabado de muro coincidente con la BD de excel y con el tipo de revit

        count+=1
   
roomsOutId = [x for x in roomsId if x not in roomsInId]
FEd = [x for x in AllEd if x not in TEd]
FP = [x for x in AllP if x not in TP]
lrooms=len(roomsOutId)

# output a model's information
message = """Proceso finalizado.

Nº de tipos de acabado no encotnrados: {lrooms}

Los tipos de acabado de los estilos de habitación aplicados a estas habitaciones: {roomsOutId} ;  no existen como tipo de acabado en Revit, modificar el excel o generar los tipos de acabado que faltan.

Los tipos de acabado no coincidentes son: {FEd}, de los estilos de habitación: {FP} """.format(lrooms=lrooms,roomsOutId=roomsOutId, FEd=FEd, FP=FP)
print(message)