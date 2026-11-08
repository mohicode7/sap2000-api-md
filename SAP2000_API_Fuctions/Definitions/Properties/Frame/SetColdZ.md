# SetColdZ

## Syntax

SapObject.SapModel.PropFrame.SetColdZ

## VB6 Procedure

Function SetColdZ(ByVal Name As String, ByVal MatProp
As String, ByVal t3 As Double, ByVal t2 As Double, ByVal Thickness As
Double, ByVal Radius As Double, ByVal LipDepth As Double, ByVal LipAngle
As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As
String = "", Optional ByVal GUID As String = "") As
Long

## Parameters

Name

The name of an existing or new frame section property.
If this is an existing property then that property is modified, otherwise,
a new property is added.

MatProp

The name of the material property for the section.

t3

The section depth. [L]

t2

The section width. [L]

Thickness

The section thickness. [L]

Radius

The corner radius, if any. [L]

LipDepth

The lip depth, if any. [L]

LipAngle

The lip angle measured from horizontal (0 <= LipAngle
<= 90). [deg]

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

This function initializes  a cold formed Z-type
frame section property. If this function is called for an existing frame
section property, all items for the section are reset to their default
value.

The function returns zero if the section property is
successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropColdZ()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, 
 MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'set new frame section property

      ret = SapModel.PropFrame.SetColdZ("CZ1", 
 Name, 10, 3.2, 0.08, 0.3, 0.6, 60)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetColdZ](GetColdZ.md)
