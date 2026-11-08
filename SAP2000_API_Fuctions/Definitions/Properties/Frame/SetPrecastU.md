# SetPrecastU

## Syntax

SapObject.SapModel.PropFrame.SetPrecastU

## VB6 Procedure

Function SetPrecastU(ByVal Name As String, ByVal MatProp As String, ByRef b() As Double, ByRef d() As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

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

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function initializes a precast concrete U girder frame section property. If this function is called for an existing frame section property, all items for the section are reset to their default value.

The function returns zero if the section property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropPrecastU()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim bb() As Double

      Dim dd() As Double



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

[GetPrecastU](GetPrecastU.md)
