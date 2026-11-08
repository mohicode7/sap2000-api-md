# SetAngle\_1

## Syntax

SapObject.SapModel.PropFrame.SetAngle\_1

## VB6 Procedure

Function SetAngle\_1(ByVal Name As String, ByVal MatProp
As String, ByVal t3 As Double, ByVal t2 As Double, ByVal tf As Double,
ByVal tw As Double, ByVal FilletRadius As Double, Optional ByVal Color
As Long = -1, Optional ByVal Notes As String = "", Optional
ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property.
If this is an existing property, that property is modified; otherwise,
a new property is added.

MatProp

The name of the material property for the section.

t3

The vertical leg depth. [L]

t2

The horizontal leg width. [L]

tf

The horizontal leg thickness. [L]

tw

The vertical leg thickness. [L]

FilletRadius

The fillet radius. [L]

Color

The display color assigned to the section. If Color
is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned
to the section. If this item is input as Default, then the program assigns
a GUID to the section.

## Remarks

This function initializes an angle-type frame section
property. If this function is called for an existing frame section property,
all items for the section are reset to their default value.

The function returns zero if the section property is
successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropAngle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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

      ret = SapModel.PropFrame.SetAngle_1("ANGLE1", 
 "A992Fy50", 6, 4, 0.5, 0.5, 0.2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 25.0.0. This function replaces
the [SetAngle](../../../Obsolete_Functions/SetAngle_%7BFrame%7D_old.md).

## See Also

[GetAngle\_1](GetAngle_%7BFrame%7D.md)
