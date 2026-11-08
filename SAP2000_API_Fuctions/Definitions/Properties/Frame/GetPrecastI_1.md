# GetPrecastI\_1

## Syntax

SapObject.SapModel.PropFrame.GetPrecastI\_1

## VB6 Procedure

Function GetPrecastI\_1(ByVal Name As String, ByRef FileName
As String, ByRef MatProp As String, ByRef b() As Double, ByRef d() As
Double, ByRef t() As Double, ByRef c As Double, ByRef Color As Long, ByRef
Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing precast concrete I girder frame
section property.

FileName

If the section property was imported from a property
file, this is the name of that file. If the section property was not imported,
this item is blank.

MatProp

The name of the material property for the section.

b

This is an array, dimensioned to 3, containing the horizontal
section dimensions. [L]

b(0) = B1 (> 0)

b(1) = B2 (> 0)

b(2) = B3 (>= 0)

b(3) = B4 (>= 0)

Section dimensions B1 through B4 are defined on the
precast concrete I girder definition form.

d

This is an array, dimensioned to 6, containing the vertical
section dimensions. [L]

d(0) = D1 (> 0)

d(1) = D2 (> 0)

d(2) = D3 (>= 0)

d(3) = D4 (>= 0)

d(4) = D5 (> 0)

d(5) = D6 (>= 0)

d(6) = D7 (>=0)

Section dimensions D1 through D7 are defined on the
precast concrete I girder definition form.

t

This is an array, dimensioned to 1, containing the web
thickness dimensions. [L]

T(0)
= T1 (> 0)

T(1) = T2 (> 0)

Section dimensions T1 and T2 are defined on the precast
I girder definition form.

c

The bottom flange chamfer dimension, denoted as C1 on
the precast concrete I girder definition form.

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned
to the section.

## Remarks

This function retrieves frame section property data
for a precast concrete I girder frame section.

The function returns zero if the section property data
is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropPrecastI_1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim MatProp As String

      Dim bb() As Double

      Dim dd() As Double
      Dim tt() As Double

      Dim b() As Double

      Dim d() As Double
      Dim t() As Double
      Dim cc As Double
      Dim c As Double

      Dim Color As Long

      Dim Notes As String

      Dim GUID As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'set new frame section property

      ReDim bb(3)

      ReDim dd(6)
      ReDim tt(1)

      bb(0) = 12

      bb(1) = 16

      bb(2) = 0

      bb(3) = 0

      dd(0) = 28

      dd(1) = 4

      dd(2) = 3

      dd(3) = 0

      dd(4) = 5

      dd(5) = 5
      dd(6) = 0
      tt(0) = 6
      tt(1) = 6
      cc = 0

      ret = SapModel.PropFrame.SetPrecastI_1("PC1", 
 "4000Psi", bb, dd, tt, cc)



   'get frame section property data

      ret = SapModel.PropFrame.GetPrecastI_1("PC1", 
 FileName, MatProp, b, d, t, c, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.2.0.

This function supersedes GetPrecast1.

## See Also

[SetPrecastI\_1](SetPrecastI_1.md)
