# Walltype-Export-Autodesk-Revit-API 

Export revit walltypes to a docx list
Program is made for architects to automize the process of creating wall type lists of the project.
It uses Autodesk Revit API and a Revit addon PyRevit.
Program uses object oriented programming from Revit API to gather data from objects existing within the Revit file.
Then it creates a new docx file with two types of wall types presentation along with wall materials and their respective widths.

This type of use saves many hours of architects manual work and adds to Revit base functionality of Building Information Modeling.


script.py <- part of the PyRevit custom toolbar. This file executes an executable that works with docx library for python.

main.py <- base for a second step executable that is opened at the end of script.py after the data is collected. This file works with docx library and the data to create a visually coherent list of walls and their properties.
