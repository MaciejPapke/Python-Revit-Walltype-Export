__doc__='Creating a wall type database'
__author__='Maciej Papke'
__title__='WallType\ndocumentation'

from Autodesk.Revit import DB
from Autodesk.Revit.DB import WallType, ElementType, FilteredElementCollector, BuiltInCategory, CompoundStructureLayer, CompoundStructure, ElementId, Element
doc = __revit__.ActiveUIDocument.Document #Gathering existing document into a variable doc
import clr
clr.AddReference("RevitNodes") #importing revit API

import os

def get_all_elements(cat): #Setting function to get all elements of given type
    cl = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsElementType()
    return cl
    #Input a BuiltInCategory.OST_{what u want here}
    #Outputs list of all types of input class elements in the project



all_el_type = get_all_elements(BuiltInCategory.OST_Walls) #Putting all elements of given type into the all_el_type variable
#[All element types]


#Deleting curtain walls through mask
x=[]
for i in all_el_type:
    x.append(i)
xfalse = []
for i in x:
    if (str(i.Kind)) != 'Basic':
        xfalse.append(1)
    else:
        xfalse.append(0)
for i in range(len(xfalse)):
    if xfalse[i] == 1:
        del x[i]

#Appending Wall types` names
all_el_names = []
for i in x:
	all_el_names.append(Element.Name.__get__(i))
#[Names of wall types]

all_el_function = []
for i in x:
	all_el_function.append(str(i.Function).replace('Autodesk.Revit.DB.WallFunction.', ''))#Appending wall type`s fuction as a string
#[Wall type function string]

comp_structures = []
for i in x:
    comp_structures.append(i.GetCompoundStructure()) #Appending a list of CompoundStructure objects of each element type
comp_structures = list(filter(None, comp_structures)) #Cleaning the list from None values
#[Compound Structures class]

comp_structures_layers = []
counter = 0
for i in comp_structures:
	comp_structures_layers.append(i.GetLayers())
#[[Compound Structures Layers class]] - list of lists

layers_width = []
for i in comp_structures_layers:
    layers_width.append(['%.2f' % round(y.Width*30.48, 2) for y in i]) #Appending lists of layer widths switched to meters
#[[Widths of layers in cm]] - list of lists

layers_materialId = []
for i in comp_structures_layers:
	layers_materialId.append([y.MaterialId for y in i]) #Appending lists of layer`s material ElementIds
#[[MaterialId class]] - list of lists

layers_materialname = []
for i in layers_materialId:
	layers_materialname.append([str(doc.GetElement(y).Name) for y in i]) #Appending a list of Material names
#[[Material names of layers]] - list of lists
"""
print(all_el_names)
print(layers_width)
print(layers_materialname)
"""


#Comma separated values
import csv
#Data
headers = [all_el_names] #list
data = [layers_width, layers_materialname]



dir= str(raw_input("Input file directory: "))
"""
#Creating a csv file and writing data
with open(dir + r"\walltype-database.csv" , 'wb') as file:
    w = csv.writer(file)
    for i in headers:
        w.writerow(i)
    for i in data:
        for y in i:
            w.writerow(y)
"""

#Documentation to txt
with open(dir + r"\walltype-listed.txt" , 'w') as file:
    file.write(str(all_el_names))
    file.write('|')
    file.write(str(layers_width))
    file.write('|')
    file.write(str(layers_materialname))

#Launching exe
os.startfile("C:\Users\macpa\Desktop\Coding\presentation\Stylize_v-2\main.exe")