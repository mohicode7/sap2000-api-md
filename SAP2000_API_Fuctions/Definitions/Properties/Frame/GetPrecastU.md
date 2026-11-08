# GetPrecastU

## Syntax

SapObject.SapModel.PropFrame.GetPrecastU

## VB6 Procedure

Function GetPrecastU(ByVal Name As String, ByRef FileName As String, ByRef MatProp As String, ByRef b() As Double, ByRef d() As Double, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing precast concrete U girder frame section property.

FileName

If the section property was imported from a property file, this is the name of that file. If the section property was not imported, this item is blank.

MatProp

The name of the material property for the section.

b

This is an array, dimensioned to 5, containing the horizontal section dimensions. [L]

b(0) = B1 (> 0)

b(1) = B2 (> 0)

b(2) = B3 (> 0)

b(3) = B4 (>= 0)

b(4) = B5 (>= 0)

b(5) = B6 (>= 0)

Section dimensions B1 through B6 are defined on the precast concrete U girder definition form.

d

This is an array, dimensioned to 6, containing the vertical section dimensions. [L]

d(0) = D1 (> 0)

d(1) = D2 (> 0)

d(2) = D3 (>= 0)

d(3) = D4 (>= 0)

d(4) = D5 (>= 0)

d(5) = D6 (>= 0)

d(6) = D7 (>= 0)

Section dimensions D1 through D7 are defined on the precast concrete U girder definition form.

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for a precast concrete U girder frame section.

The function returns zero if the section property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropPrecastU()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim MatProp As String

      Dim bb() As Double

      Dim dd() As Double

      Dim b() As Double

      Dim d() As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set present units to kN-mm

      ret = SapModel.SetPresentUnits(kN_mm_C)



   'set new frame section property

      ReDim bb(5)

      ReDim dd(6)

      bb(0) = 2350

      bb(1) = 1500

      bb(2) = 200

      bb(3) = 75

      bb(4) = 143.8447

      bb(5) = 0

      dd(0) = 1700

      dd(1) = 175

      dd(2) = 75

      dd(3) = 0

      dd(4) = 0

      dd(5) = 125

      dd(6) = 175

      ret = SapModel.PropFrame.SetPrecastU("PC1", "4000Psi", bb, dd)



   'get frame section property data

      ret = SapModel.PropFrame.GetPrecastU("PC1", FileName, MatProp, b, d, Color, Notes, GUID)



   'set present units to kip-in

      ret = SapModel.SetPresentUnits(kip_in_F)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetPrecastU](SetPrecastU.md)
