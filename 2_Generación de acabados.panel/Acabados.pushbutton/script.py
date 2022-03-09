# -*- coding: utf-8 -*-

#pyRevit info
__title__ = """Generar 
acabados"""
__doc__ = ''
__author__  = 'ADT'

#_________________iMPORTACIONES_______________

#Importar RevitAPI
import clr
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import*
from Autodesk.Revit.UI import *
from Autodesk.Revit.ApplicationServices import *
from Autodesk.Revit.Attributes import *

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
from RevitServices.Transactions import TransactionManager
    
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
doc1 = DocumentManager.Instance.CurrentDBDocument

clr.AddReference("System")
from System.Collections.Generic import List

import itertools

#_________________UTILIDADES_______________

#Conversión de unidades
ft = 304.8 # 1ft = 304.8mm

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

#PARÁMETRO 2
#with forms.WarningBar(title='Selecciona el parámetro de CONSULTA donde está el código de acabado en el suelo'):
#    selected_parameters = forms.select_parameters(floor)[0]
#    Nombre_parametro_consulta = selected_parameters.name
    
    
#Nivel de las habitaciones
levelrooms = []
names = []

for room in rooms:
    LevelId = room.Level.Id
    Level = room.Level
    levelrooms.append(LevelId)

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

#___________________RELLENAR PARÁMETROS DE REVIT CON LA BASE DE DATOS DE EXCEL_______________
t = Transaction(doc,"Creación de suelos")
t.Start()

count=0
count1=0
for r, Ed in [(r,Ed) for r in rooms for Ed in EData]:
    
    if Ed[0]==(rooms[count].LookupParameter(Nombre_parametro_a_consultar).AsString()):

        #Param Acabado de muro
        WallParam = rooms[count].LookupParameter('Acabado de muro')
        WallParam.Set(Ed[1])
        
        #Param Acabado de suelo
        WallParam = rooms[count].LookupParameter('Acabado del suelo')
        WallParam.Set(Ed[2])
        
        #Param Acabado de techo
        WallParam = rooms[count].LookupParameter('Acabado del techo')
        WallParam.Set(Ed[3])
        
    
    count1 +=1
    
    if len(EData)==count1:
        count +=1
        count1=0

t.Commit() 

#_________________iNTERACCIÓN 3_SELECCIÓN DE TIPO DE ACABADO_______________

Tipo= []
TId= []
TipoType = []
tipos = []

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

        for t in types:
            if ATS0 == t.LookupParameter(Nombre_parametro_consulta).AsString():
                TId.append(t.Id)
                types1.append(t)
                rooms2.append(rooms1[count])
                Lrooms2.append(rooms1[count].Level.Id)
                
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
    
    #Busqueda de los nombres de estilo de excel en los parametros de tipo
    for r, Ed in [(r,Ed) for r in rooms for Ed in EData]:
        
        if Ed[0]== r.LookupParameter(Nombre_parametro_a_consultar).AsString():
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
    
    #1. Rooms seleccionadas + todos los tipos de acabado
    #2. Base de datos filtrada  + habitaciones filtradas: coincidentes como estilo de habitacion
    #3. Tipos de muro filtrados + habitacion filtradas a la par: coincidentes con estilo de habitacion y con tipo de muro
    #4. Filtrado en funcion del elemento vecino. filtrar a la par: tipos de acabadp, nivel de habitación y altura de muro
    #5. Offset de crvs con la mitad del ancho de muro del tipo de acabado
    #6. Wall.Create(curvas(for segment in curves), Tipo de acabado, nivel de habitación, Altura de muro)
    
    #Buscar las coincidencias entre la base de datos de excel con el estilo de habitacion en revit
    for r, Ed in [(r,Ed) for r in rooms for Ed in EData]:
        if Ed[0]== r.LookupParameter('Departamento').AsString():
            rooms1.append(r) #Habitaciones existentes como estilo de habitacion
            EData1.append(Ed) #Base de datos de excel filtrada
            
    #Buscar las coincidencias entre la base de datos de excel filtrada en el paso anterior y los tipos de acabados creados en revit
    types1=[]
    AWS = []
    
    count=0
    for Ed in EData1:
        AWS0 = (Ed[1].split("/"))[0]
        
        AWS.append(AWS0)
        
        for t in types:
            if AWS0 == t.LookupParameter(Nombre_parametro_consulta).AsString():
                
                rooms2.append(rooms1[count]) #Habitaciones filtradas a la par con los tipos
                Lrooms.append(rooms1[count].Level) #Niveles de habitación existentes como estilo de habitación

                TId.append(t.Id) #Id del tipo de muro coincidente con la BD de excel y con el tipo de revit
                types1.append(t) #Tipo de acabado de muro coincidente con la BD de excel y con el tipo de revit
                ValueWidth.append(t.Width) #Ancho de acabado de muro coincidente con la BD de excel y con el tipo de revit

                
        count+=1


#_________________OBTENCIÓN DE CONTORNOS DE HABITACIÓN_______________

calculator = SpatialElementGeometryCalculator(doc)
options = Autodesk.Revit.DB.SpatialElementBoundaryOptions()
boundloc = Autodesk.Revit.DB.AreaVolumeSettings.GetAreaVolumeSettings(doc).GetSpatialElementBoundaryLocation(SpatialElementType.Room)
options.SpatialElementBoundaryLocation = boundloc
curve_arrays = []
ca = []

for r in rooms:
    
    for group in r.GetBoundarySegments(options):
        l=[]
        
        for segment in group:
            l.append(segment.GetCurve())
        curve_arrays.append(l)


#_________________GENERACIÓN DE ACABADOS_______________

#_________________ACABADOS DE TECHO
if Acabados == 'Acabado de techo':
    c = []
    loop=[]
    loops=[]
    cens=[]
    Height=[]
    
    Af=[]
    AId=[]
    CodA=set()
    RId=set()
    
    count = 0
        
    calculator = SpatialElementGeometryCalculator(doc)
    for r in rooms2:
        results = calculator.CalculateSpatialElementGeometry(r)
        Solid = results.GetGeometry()
        Height.append(r.UnboundedHeight)
        
        for f in Solid.Faces:
            uv=UV(0.5,0.5)
            normal=f.ComputeNormal(uv).ToString()
            
            if normal == '(0.000000000, 0.000000000, 1.000000000)':
                faces=results.GetBoundaryFaceInfo(f)
                crv=f.GetEdgesAsCurveLoops()
                
                t = Transaction(doc,"Creación de techos")
                t.Start()
                
                ceiling=Ceiling.Create(doc, crv, TId[count], Lrooms2[count])
                param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
                param.Set(Height[count])
                t.Commit()  
                
                Af.append(ceiling)
                AId.append(ceiling.Id.ToString())
                TypId=doc.GetElement(TId[count])
                CodA.add(TypId.LookupParameter(Nombre_parametro_consulta).AsString())
                RId.add(r.Id.ToString()) 
                
        count += 1
    lAf=len(Af)

#_________________ACABADOS DE SUELO
elif Acabados == 'Acabado de suelo':
    c = []
    loop=[]
    loops=[]
    cens=[]
    Af=[]
    AId=[]
    CodA=set()
    RId=set()
    
    count= 0
    
    t = Transaction(doc,"Creación de suelos")
    t.Start()    
       
    calculator = SpatialElementGeometryCalculator(doc)
    
    for r in rooms2:
        results = calculator.CalculateSpatialElementGeometry(r)
        Solid = results.GetGeometry()

        for f in Solid.Faces:
            uv=UV(0.5,0.5)
            normal=f.ComputeNormal(uv).ToString()
            
            if normal == '(0.000000000, 0.000000000, -1.000000000)':
                faces=results.GetBoundaryFaceInfo(f)
                crv=f.GetEdgesAsCurveLoops()
                
                ceiling=Floor.Create(doc, crv, TId[count], Lrooms2[count])
                Af.append(ceiling)
                AId.append(ceiling.Id.ToString())
                TypId=doc.GetElement(TId[count])
                CodA.add(TypId.LookupParameter(Nombre_parametro_consulta).AsString())
                RId.add(r.Id.ToString())
        count += 1
        
    t.Commit()
    lAf=len(Af)
#_________________ACABADOS DE MURO
elif Acabados == 'Acabado de muro':

    calculator = SpatialElementGeometryCalculator(doc)
    options = Autodesk.Revit.DB.SpatialElementBoundaryOptions()
    boundloc = Autodesk.Revit.DB.AreaVolumeSettings.GetAreaVolumeSettings(doc).GetSpatialElementBoundaryLocation(SpatialElementType.Room)
    options.SpatialElementBoundaryLocation = boundloc

    curve_arrays = []
    TId2=[]
    Wwidth2=[]
    wHeights = []
    levelrooms1 = []
    roomElems = []
    LevelId=[]
    rooms3=[]     
    

    loops=[]
    loop=[]
    count=0

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
            rooms3T=[] 
            
            for segment in group:
                
                if doc.GetElement(segment.ElementId) is None:
                    roomElemsT.append(None)
                    
                    lT.append(segment.GetCurve())
                    TId2T.append(TId[count])
                    Wwidth2T.append(types1[count].Width)
                    LevelIdT = rooms2[count].Level.Id
                    levelrooms1T.append(LevelIdT)
                    wHeightsT.append(rooms2[count].UnboundedHeight)
                    rooms3T.append(r)
                    
                elif doc.GetElement(segment.ElementId).Category.Id.ToString() ==  "-2001352":
                    roomElemsT.append(doc.GetElement(segment.ElementId))
                    
                    lT.append(segment.GetCurve())
                    TId2T.append(TId[count])
                    Wwidth2T.append(types1[count].Width)
                    LevelIdT = rooms2[count].Level.Id
                    levelrooms1T.append(LevelIdT)
                    wHeightsT.append(rooms2[count].UnboundedHeight)
                    rooms3T.append(r)
                    
                elif doc.GetElement(segment.ElementId).Category.Id.ToString() ==  "-2000066":
                    roomElemsT.append(doc.GetElement(segment.ElementId))
                    
                    lT.append(segment.GetCurve())
                    TId2T.append(TId[count])
                    Wwidth2T.append(types1[count].Width)
                    LevelIdT = rooms2[count].Level.Id
                    levelrooms1T.append(LevelIdT)
                    wHeightsT.append(rooms2[count].UnboundedHeight)
                    rooms3T.append(r)
                    
                else:                    
                    tipoel=doc.GetElement(doc.GetElement(segment.ElementId).GetTypeId())
                    CodTipo = tipoel.LookupParameter(Nombre_parametro_consulta).AsString()
                    
                    
                    if CodTipo != None and CodTipo[0] == AWS[count][0]: #If si existe acabado (R)
                        
                         if CodTipo != AWS[count]: #If si no coincide: hay que eliminarlo para generarlo a posteriori

                            t = Transaction(doc,"Creación de suelos")
                            t.Start()    
                            doc.Delete(segment.ElementId)
                            t.Commit() 
                            lT.append(segment.GetCurve())
                            TId2T.append(TId[count])
                            Wwidth2T.append(types1[count].Width)
                            LevelIdT = rooms2[count].Level.Id
                            levelrooms1T.append(LevelIdT)
                            wHeightsT.append(rooms2[count].UnboundedHeight)
                            roomElemsT.append(doc.GetElement(segment.ElementId)) 
                            rooms3T.append(r)
                                                     
                    else: #else si no existe y hay que generarlo                        
                        lT.append(segment.GetCurve())
                        TId2T.append(TId[count])
                        Wwidth2T.append(types1[count].Width)
                        LevelIdT=(rooms2[count].Level.Id)
                        levelrooms1T.append(LevelIdT)
                        wHeightsT.append(rooms2[count].UnboundedHeight)
                        roomElemsT.append(doc.GetElement(segment.ElementId))
                        rooms3T.append(r)
                        
            #ordenar listas en sublistas de elementos, cada sublista corresponde a un loop de contornos            
            curve_arrays.append(lT)
            roomElems.append(roomElemsT)
            TId2.append(TId2T)
            Wwidth2.append(Wwidth2T)
            LevelId.append(LevelIT)
            levelrooms1.append(levelrooms1T)
            wHeights.append(wHeightsT)
            rooms3.append(rooms3T)
            
        count+=1
    
    #Discriminar vecinos en función del elemento
    t = Transaction(doc,"Acabados de muro")
    t.Start()

    ie =[]
    #final_curve_arrays=[]
    
    count = 0
    
    i1=[]
    for i in range(len(roomElems)):
        f_curve_arrays=[]
        j1=[]
        
        #Filtrar los elementos vecinos de los limites de la habitación en funcion de su categoria, para discriminar lo que son separadores de habitación o muros de cortina
        for j in range(len(roomElems[i])):
            
            if roomElems[i][j] is not None and roomElems[i][j].Category.Id.ToString() == "-2000011" and roomElems[i][j].WallType.Kind != WallKind.Curtain:
                j1.append(j)
                
            if roomElems[i][j] is None:
                j1.append(j)

            if roomElems[i][j] is not None and roomElems[i][j].Category.Id.ToString() == "-2001352":
                j1.append(j)
                
        i1.append(j1)
    
    t.Commit() 
    
    #Crear muros iniciales
    t = Transaction(doc,"Creación de suelos")
    t.Start()

    
    LrebldCrv=[]
    TId3=[]
    levelrooms2=[]
    walls=[]
    for i in range(len(i1)):
        ws = []
        for j in i1[i]:
            
            #Volver a dibujar los segmentos con los que se generan los muros   
            if XYZ.IsAlmostEqualTo(curve_arrays[i][j].ComputeDerivatives(0,True).BasisX,curve_arrays[i][j].ComputeDerivatives(1,True).BasisX):
                rebldCrv = Line.CreateBound(curve_arrays[i][j].Evaluate(0,True), curve_arrays[i][j].Evaluate(1,True))
                
            else:
                rebldCrv = Arc.Create(curve_arrays[i][j].Evaluate(0,True),  curve_arrays[i][j].Evaluate(0.5,True), curve_arrays[i][j].Evaluate(1,True))
                
            LrebldCrv.append(rebldCrv)
            w = Wall.Create(doc, rebldCrv, TId2[i][j], levelrooms1[i][j], 10, 0, False, False)
            
            ws.append(w)
            TId3.append(TId2[i][j])
            levelrooms2.append(levelrooms1[i][j])
            
        walls.append(ws)

    t.Commit()
        
    #Generar los muros finales
    Offsetedwalls = []
    Onewlocation = []

    t = Transaction(doc,"Acabados de muro")
    t.Start()
    
    count = 0
    
    wHeights1=[]
    newlocation=[]

    for i in range(len(walls)):
        newlocationT=[]
        for j in range(len(walls[i])):
            
            cv= walls[i][j].Location
            newlocationT.append(cv.Curve.CreateTransformed(Transform.CreateTranslation(walls[i][j].Orientation * (Wwidth2[i][i1[i][j]]/2))))
            doc.Delete(walls[i][j].Id)
        newlocation.append(newlocationT)

    t.Commit()
    
    
    #_______________AJUSTAR LONGITUDES DE MUROS PARA ENCAJAR BIEN LAS ESQUINAS_______________
    
    #Ajuste de los muros hacia un sentido
    t = Transaction(doc,"Acabados de muro")
    t.Start()
    
    cnewlocation=[]
    resultArray = clr.Reference[DB.IntersectionResultArray]()
 
    for j in range(0, (len(newlocation))):
        cnewlocationT=[]
    
        for i in range(len(newlocation[j])):
            
            if i != (len(newlocation[j])-1): #Intersección entre un segmento y el siguiente
                result = newlocation[j][i].Intersect(newlocation[j][i+1], resultArray).ToString()

                if result == 'Overlap': #Si el contorno interseca con el siguiente
                    pstart = newlocation[j][i].GetEndParameter(0)
                    pend = newlocation[j][i].GetEndParameter(1)
                    newlocation[j][i].MakeUnbound()
                    newlocation[j][i].MakeBound(pstart, pend -(Wwidth2[j][i]*(0.666)))
                    cnewlocationT.append(newlocation[j][i])
                    
                elif result == 'Subset': #Si el contorno es paralelo con el siguiente
                    cnewlocationT.append(newlocation[j][i])
                     
                else: #Si el contorno no interseca ni es paralelo con el siguiente
                    pstart = newlocation[j][i].GetEndParameter(0)
                    pend = newlocation[j][i].GetEndParameter(1)
                    newlocation[j][i].MakeUnbound()
                    newlocation[j][i].MakeBound(pstart, pend +(Wwidth2[j][i]*(0.666)))
                    cnewlocationT.append(newlocation[j][i])
                    
            else: #Intersección entre el último segmento y el primero de la lista
                result = newlocation[j][i].Intersect(newlocation[j][0], resultArray).ToString()

                if result == 'Overlap': #Si el contorno interseca con el primero
                    pstart = newlocation[j][i].GetEndParameter(0)
                    pend = newlocation[j][i].GetEndParameter(1) 
                    newlocation[j][i].MakeUnbound()
                    newlocation[j][i].MakeBound(pstart, pend -(Wwidth2[j][i]*(0.666)))
                    cnewlocationT.append(newlocation[j][i])
                
                elif result == 'Subset': #Si el contorno es paralelo con el primero
                    cnewlocationT.append(newlocation[j][i])
                    
                else: #Si el contorno no interseca con el primero
                    pstart = newlocation[j][i].GetEndParameter(0)
                    pend = newlocation[j][i].GetEndParameter(1)
                    newlocation[j][i].MakeUnbound()
                    newlocation[j][i].MakeBound(pstart, pend +(Wwidth2[j][i]*(0.666)))
                    cnewlocationT.append(newlocation[j][i])     
            
        cnewlocation.append(cnewlocationT)
    
    #Ajuste de los muros hacia el otro sentido
    
    #Dar la vuelta a las listas
    cnewlocation=[elem[::-1] for elem in cnewlocation]
    newlocation=[elem[::-1] for elem in newlocation]
    Wwidth2=[elem[::-1] for elem in Wwidth2]
    
    for j in range(0, (len(newlocation))):
        cnewlocationT=[]

        for i in range(len(newlocation[j])):
            
            if i != (len(newlocation[j])-1):  #Intersección entre un segmento y el anterior
                result = newlocation[j][i].Intersect(newlocation[j][i+1], resultArray).ToString()

                if result == 'Overlap': #Si el contorno interseca con el anterior
                    pstart = newlocation[j][i].GetEndParameter(0)
                    pend = newlocation[j][i].GetEndParameter(1)
                    newlocation[j][i].MakeUnbound()
                    newlocation[j][i].MakeBound(pstart+(Wwidth2[j][i]*(0.666)), pend )
                    cnewlocationT.append(newlocation[j][i])
					
                elif result == 'Subset': #Si el contorno es paralelo con el anterior
                    cnewlocationT.append(newlocation[j][i])
					
                else: #Si el contorno no interseca ni es paralelo con el anterior
                    pstart = newlocation[j][i].GetEndParameter(0)
                    pend = newlocation[j][i].GetEndParameter(1)
                    newlocation[j][i].MakeUnbound()
                    newlocation[j][i].MakeBound(pstart-(Wwidth2[j][i]*(0.666)), pend)
                    cnewlocationT.append(newlocation[j][i])
                    
            else: #Intersección entre el último segmento y el primero de la lista
                result = newlocation[j][i].Intersect(newlocation[j][0], resultArray).ToString()

                if result == 'Overlap': #Si el contorno interseca con el último
                    pstart = newlocation[j][i].GetEndParameter(0)
                    pend = newlocation[j][i].GetEndParameter(1)
                    newlocation[j][i].MakeUnbound()
                    newlocation[j][i].MakeBound(pstart+(Wwidth2[j][i]*(0.666)), pend)
                    cnewlocationT.append(newlocation[j][i])
					
                elif result == 'Subset': #Si el contorno es paralelo con el último
                    cnewlocationT.append(newlocation[j][i])
                
                else: #Si el contorno no interseca con el último
                    pstart = newlocation[j][i].GetEndParameter(0)
                    pend = newlocation[j][i].GetEndParameter(1)
                    newlocation[j][i].MakeUnbound()
                    newlocation[j][i].MakeBound(pstart-(Wwidth2[j][i]*(0.666)), pend)
                    cnewlocationT.append(newlocation[j][i])     

        cnewlocation[j]=cnewlocationT

    #Dar la vuelta a las listas
    cnewlocation=[elem[::-1] for elem in cnewlocation]
    newlocation=[elem[::-1] for elem in newlocation]
    Wwidth2=[elem[::-1] for elem in Wwidth2]

    t.Commit()  
    
    Af=[]
    AId=[]
    CodA=set()
    RId=set()
    
    #Generación de muros
    t = Transaction(doc,"Acabados de muro")
    t.Start()
    
    for i in range(len(i1)):
        count=0
        OffsetedwallsT=[]
        for j in i1[i]:

            wall1 = Wall.Create(doc, cnewlocation[i][count], TId2[i][j], levelrooms1[i][j], wHeights[i][j], 0, False, False)
            
            OffsetedwallsT.append(wall1)
            AId.append(wall1.Id.ToString())
            TypId=doc.GetElement(TId2[i][j])
            CodA.add(TypId.LookupParameter(Nombre_parametro_consulta).AsString())
            RId.add(rooms3[i][j].Id.ToString())
            
            count +=1
            
        Offsetedwalls.append(OffsetedwallsT)
    t.Commit()    
    
    lAf=len(OffsetedwallsT)
    
    #Unión de muros con sus vecinos para abrir los huecos
    t = Transaction(doc,"Acabados de muro")
    t.Start()
    
    count=0

    for i in range(len(i1)):
        count=0
        for j in i1[i]:

            if roomElems[i][j] is not None and roomElems[i][j].Category.Id.ToString() == "-2000011" and len(roomElems[i][j].FindInserts(True, True, True, True)) != 0:
                JoinGeometryUtils.JoinGeometry(doc,Offsetedwalls[i][count],roomElems[i][j])

            count +=1

    t.Commit()
    
# output a model's information
message = """Proceso finalizado.
Número de acabados generados: {lAf}.
Habitaciones en las que se ha generado acabado: {RId}.
Acabados generados:{AId}.
Codigos de acabados generados: {CodA}. """.format(CodA=CodA,AId=AId,lAf=lAf,RId=RId)

print(message)