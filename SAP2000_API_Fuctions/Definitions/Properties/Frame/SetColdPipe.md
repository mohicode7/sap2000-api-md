# SetColdPipe

## Syntax

SapObject.SapModel.PropFrame.SetColdPipe

## VB6 Procedure

Function SetColdPipe(ByVal Name As String, ByVal MatProp
As String, ByVal t3 As Double, ByVal Thickness As Double, Optional ByVal
Color As Long = -1, Optional ByVal Notes As String = "", Optional
ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property.
If this is an existing property, that property is modified; otherwise,
a new property is added.

MatProp

The name of the material property for the section.

t3

The section depth. [L]

Thickness

The section thickness. [L]

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

This function initializes a cold formed pipe frame section
property. If this function is called for an existing frame section property,
all items for the section are reset to their default value.

The function returns zero if the section property is
successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropColdPipe()

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

      ret = SapModel.PropFrame.SetColdPipe("CP1", 
 Name, 8.0, 0.08)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 22.2.0.

## See Also

[GetColdPipe](GetColdPipe.md)
