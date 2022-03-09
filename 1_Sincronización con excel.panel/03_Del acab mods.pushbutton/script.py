# -*- coding: utf-8 -*-

#pyRevit info
__title__ = """Eliminar 
acabados modificados"""
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
          
    calculator = SpatialElementGeometryCalculator(doc)  
    caras_superiores=[] 
    gs=[]
    techos=[]
    rooms_con_techo=[]
    mallas_finales=[]
    f_ceilings=[]
    for c in ceilings:
        gs.append(c.get_Geometry(Options()))
        
    #Obtenemos una lista con los sólidos de cada suelo, aunque no es lo normal, podría tener más de uno:
    ceilings_f=[]
    mallas=[]
    
    count=0
    for solidos in gs:
        for solido in solidos:
            caras= solido.Faces
            for cara in caras:
                if isinstance(cara,PlanarFace):
                    
                    #Nos quedamos con caras con pendiente inferior al 15%
                    if cara.FaceNormal[2] > 0.988:
                        malla = cara.Triangulate(0)
                        mallas.append(malla)
                        f_ceilings.append(ceilings[count])
                        break
        count+=1

    # Con las mallas, códigos y suelos procedemos a inscribir los parámetros de los códigos en cada habitación.
    t = Transaction(doc, "Chequeo Acabados, errores")
    t.Start()
    RId=[]
    
    for m in mallas:
        #Iteramos tantas veces como triángulos haya en la malla
        for idx in range(m.NumTriangles):
            
            #Vamos sacando baricentros de triángulos hasta que alguno esté en una room, hallamos la habitación donde cae el triángulo.
            triangulo = m.Triangle[idx]
            
            a, b, c = triangulo.Vertex[0], triangulo.Vertex[1], triangulo.Vertex[2]
            
            bar = XYZ((a.X+b.X+c.X)/3,(a.Y+b.Y+c.Y)/3,(a.Z+b.Z+c.Z)/3 -3)
                
        rooms_con_techo.append(doc.GetRoomAtPoint(bar))
        
    final_rooms=[]
    final_ceilings=[]
    count=0
    for r in rooms_con_techo:
        if r is not None:
            final_rooms.append(r)
            final_ceilings.append(f_ceilings[count])
        count+=1

    par=[]
    for f in final_ceilings:
        tipo_id = f.GetTypeId() 
        tipo = doc.GetElement(tipo_id) 
        descripcion = tipo.LookupParameter(Nombre_parametro_consulta) 
        par.append(descripcion.AsString())
        
    ACod=[]
    ARCod=[]
    la=[]
    for i in range(len(final_rooms)):   
        if  par[i][0] == 'T':
            Pr=(final_rooms[i].LookupParameter('Acabado del techo').AsString()).split("/") 
            Pr0=Pr[0]      
            if par[i] != Pr0:
                doc.Delete(final_ceilings[i].Id)
                ACod.append(par[i])
                ARCod.append(Pr0)
                RId.append(final_rooms[i].Id.ToString()) 
    la=len(ACod)
    t.Commit()

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
        
    calculator = SpatialElementGeometryCalculator(doc)  
    caras_superiores=[] 
    gs=[]
    techos=[]
    rooms_con_techo=[]
    mallas_finales=[]
    f_ceilings=[]
    for c in floors:
        gs.append(c.get_Geometry(Options()))
        
    #Obtenemos una lista con los sólidos de cada suelo, aunque no es lo normal, podría tener más de uno:
    ceilings_f=[]
    mallas=[]
    
    count=0
    for solidos in gs:
        for solido in solidos:
            caras= solido.Faces
            for cara in caras:
                if isinstance(cara,PlanarFace):
                    
                    #Nos quedamos con caras con pendiente inferior al 15%
                    if cara.FaceNormal[2] > 0.988:
                        malla = cara.Triangulate(0)
                        mallas.append(malla)
                        f_ceilings.append(floors[count])
                        break
        count+=1

    # Con las mallas, códigos y suelos procedemos a inscribir los parámetros de los códigos en cada habitación.
    t = Transaction(doc, "Chequeo Acabados, errores")
    t.Start()
    RId=[]
    
    for m in mallas:
        #Iteramos tantas veces como triángulos haya en la malla
        for idx in range(m.NumTriangles):
            
            #Vamos sacando baricentros de triángulos hasta que alguno esté en una room, hallamos la habitación donde cae el triángulo.
            triangulo = m.Triangle[idx]
            
            a, b, c = triangulo.Vertex[0], triangulo.Vertex[1], triangulo.Vertex[2]
            
            bar = XYZ((a.X+b.X+c.X)/3,(a.Y+b.Y+c.Y)/3,(a.Z+b.Z+c.Z)/3 +3)
                
        rooms_con_techo.append(doc.GetRoomAtPoint(bar))
        
    final_rooms=[]
    final_ceilings=[]
    count=0
    for r in rooms_con_techo:
        if r is not None:
            final_rooms.append(r)
            final_ceilings.append(f_ceilings[count])
        count+=1

    par=[]
    for f in final_ceilings:
        tipo_id = f.GetTypeId() 
        tipo = doc.GetElement(tipo_id)
        if tipo.LookupParameter(Nombre_parametro_consulta) != None:
            descripcion = tipo.LookupParameter(Nombre_parametro_consulta) 
            par.append(descripcion.AsString())

    ACod=[]
    ARCod=[]
    la=[]
    
    for i in range(len(final_rooms)):  
        if  par[i]!=None and par[i][0] == 'P'and "_" in par[i]:
            Pr=(final_rooms[i].LookupParameter('Acabado del suelo').AsString()).split("/") 
            Pr0=Pr[0]     
            if par[i] != Pr0:
                doc.Delete(final_ceilings[i].Id)
                ACod.append(par[i])
                ARCod.append(Pr0)
                RId.append(final_rooms[i].Id.ToString()) 
    la=len(ACod)
    t.Commit()
    

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
    AWS0A=[]
    
    count=0
    for Ed in EData1:
        AWS = Ed[1].split("/")      
        AWS0 = AWS[0] 
        AWS0A.append(AWS0)
        for t in types:
            if AWS0 == t.LookupParameter(Nombre_parametro_consulta).AsString():

                rooms2.append(rooms1[count]) #Habitaciones filtradas a la par con los tipos
                Lrooms.append(rooms1[count].Level) #Niveles de habitación existentes como estilo de habitación

                TId.append(t.Id) #Id del tipo de muro coincidente con la BD de excel y con el tipo de revit
                types1.append(t) #Tipo de acabado de muro coincidente con la BD de excel y con el tipo de revit
                ValueWidth.append(t.Width) #Ancho de acabado de muro coincidente con la BD de excel y con el tipo de revit

        count+=1


    calculator = SpatialElementGeometryCalculator(doc)
    options = Autodesk.Revit.DB.SpatialElementBoundaryOptions()
    boundloc = Autodesk.Revit.DB.AreaVolumeSettings.GetAreaVolumeSettings(doc).GetSpatialElementBoundaryLocation(SpatialElementType.Room)
    options.SpatialElementBoundaryLocation = boundloc
    count=0

    ACod=[]
    ARCod=[]
    la=[]
    RId=[]
    for r in rooms2:
        
        for group in r.GetBoundarySegments(options):
            l=[]
            roomElemsT = []
            TId2T=[]
            Wwidth2T=[]
            LevelIT=[]
            levelrooms1T=[]
            wHeightsT=[]
            lT=[]
            
            for segment in group:
                
                if doc.GetElement(segment.ElementId) is not None:
                    
                    if doc.GetElement(segment.ElementId).Category.Id.ToString() ==  "-2000011":
                        tipoel=doc.GetElement(doc.GetElement(segment.ElementId).GetTypeId())
                        CodTipo = tipoel.LookupParameter(Nombre_parametro_consulta).AsString()
                        
                        if CodTipo != None and CodTipo[0] == AWS0A[count][0]: #If si existe acabado (R)
                            if CodTipo != AWS0A[count]: #If si no coincide: hay que eliminarlo para generarlo a posteriori
                                t = Transaction(doc,"Creación de suelos")
                                t.Start()    
                                doc.Delete(segment.ElementId)
                                t.Commit() 
                                ACod.append(CodTipo)
                                ARCod.append(AWS0A[count])
                                RId.append(r.Id.ToString()) 
    
        count+=1
    la=len(ACod)
    
# output a model's information
message = """Proceso finalizado.
Número de acabados eliminados: {la}.
Habitaciones con acabados inconsistentes: {RId}.
Las habitaciones tenian modelados acabados con los siguientes codigos: {ACod}.
Segun la información del codigo que aparece en Revit como estilo de habitación, deberían de ser: {ARCod}. """.format(ACod=ACod,ARCod=ARCod,la=la,RId=RId)
print(message)