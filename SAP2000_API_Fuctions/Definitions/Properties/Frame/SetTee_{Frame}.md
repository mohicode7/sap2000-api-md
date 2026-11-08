# SetTee\_1

## Syntax

SapObject.SapModel.PropFrame.SetTee\_1

## VB6 Procedure

Function SetTee\_1(ByVal Name As String, ByVal MatProp
As String, ByVal t3 As Double, ByVal t2 As Double, ByVal tf As Double,
ByVal tw As Double, ByVal FilletRadius As Double, ByVal MirrorAbout3 As
Boolean, Optional ByVal Color As Long = -1, Optional ByVal Notes As String
= "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property.
If this is an existing property, that property is modified; otherwise,
a new property is added.

MatProp

The name of the material property for the section.

t3

The section depth. [L]

t2

The flange width. [L]

tf

The flange thickness. [L]

tw

The web thickness. [L]

FilletRadius

The fillet radius. [L]

MirrorAbout3 (not applicable)

Indicates whether the section is mirrored about the
local 3-axis.

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

This function initializes a tee-type frame section property.
If this function is called for an existing frame section property, all
items for the section are reset to their default value.

The function returns zero if the section property is
successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropTee()

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

      ret = SapModel.PropFrame.SetTee_1("TEE1", 
 "A992Fy50", 12, 10, 0.6, 0.3, 0.2, False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 25.0.0. This function replaces
the [SetTee](../../../Obsolete_Functions/SetTee_%7BFrame%7D_old.md).

## See Also

[GetTee\_1](GetTee_%7BFrame%7D.md)

[SetRebarBeam](SetRebarBeam.md)

[GetRebarBeam](GetRebarBeam.md)
