# SetPrecastI (Note: Newer Function Available)

## Syntax

SapObject.SapModel.PropFrame.SetPrecastI

## VB6 Procedure

Function SetPrecastI(ByVal Name As String, ByVal MatProp As String, ByRef b() As Double, ByRef d() As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

MatProp

The name of the material property for the section.

b

This is an array, dimensioned to 3, containing the horizontal section dimensions. [L]

b(0) = B1 (> 0)

b(1) = B2 (> 0)

b(2) = B3 (> 0)

b(3) = B4 (>= 0)

Section dimensions B1 through B4 are defined on the precast concrete I girder definition form.

d

This is an array, dimensioned to 5, containing the vertical section dimensions. [L]

d(0) = D1 (> 0)

d(1) = D2 (> 0)

d(2) = D3 (>= 0)

d(3) = D4 (>= 0)

d(4) = D5 (>= 0)

d(5) = D6 (> 0)

Section dimensions D1 through D6 are defined on the precast concrete I girder definition form.

Color

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function initializes a precast concrete I girder frame section property. If this function is called for an existing frame section property, all items for the section are reset to their default value.

The function returns zero if the section property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropPrecastI()

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



   'set new frame section property

      ReDim bb(3)

      ReDim dd(5)

      bb(0) = 16

      bb(1) = 22

      bb(2) = 7

      bb(3) = 0

      dd(0) = 45

      dd(1) = 7

      dd(2) = 4.5

      dd(3) = 0

      dd(4) = 7.5

      dd(5) = 7

      ret = SapModel.PropFrame.SetPrecastI("PC1", "4000Psi", bb, dd)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

The function is obsolete and has been superseded by SetPrecastI\_1 as of version 17.2.0. This function is maintained for backward compatibility.

## See Also

[GetPrecastI](GetPrecastI.md)
