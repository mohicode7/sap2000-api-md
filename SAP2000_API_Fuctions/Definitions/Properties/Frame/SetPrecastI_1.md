# SetPrecastI\_1

## Syntax

SapObject.SapModel.PropFrame.SetPrecastI\_1

## VB6 Procedure

Function SetPrecastI\_1(ByVal Name As String, ByVal MatProp
As String, ByRef b() As Double, ByRef d() As Double, ByRef t() As Double,
ByRef c As Double,Optional ByVal Color As Long = -1, Optional ByVal Notes
As String = "", Optional ByVal GUID As String = "")
As Long

## Parameters

Name

The name of an existing or new frame section property.
If this is an existing property, that property is modified; otherwise,
a new property is added.

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

The display color assigned to the section. If Color
is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned
to the section. If this item is input as Default, the program assigns
a GUID to the section.

## Remarks

This function initializes a precast concrete I girder
frame section property. If this function is called for an existing frame
section property, all items for the section are reset to their default
value.

The function returns zero if the section property is
successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropPrecastI_1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim bb() As Double

      Dim dd() As Double
      Dim tt() As Double
      Dim cc As Double



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.2.0.

The function supersedes by SetPrecastI.

## See Also

[GetPrecastI\_1](GetPrecastI_1.md)
