# SetTube (Note: Newer function available)

## Syntax

SapObject.SapModel.PropFrame.SetTube

## VB6 Procedure

Function SetTube(ByVal Name As String, ByVal MatProp
As String, ByVal t3 As Double, ByVal t2 As Double, ByVal tf As Double,
ByVal tw As Double, Optional ByVal Color As Long = -1, Optional ByVal
Notes As String = "", Optional ByVal GUID As String = "")
As Long

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

The section width. [L]

tf

The flange thickness. [L]

tw

The web thickness. [L]

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

This function initializes a tube-type frame section
property. If this function is called for an existing frame section property,
all items for the section are reset to their default value.

The function returns zero if the section property is
successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropTube()

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

      ret = SapModel.PropFrame.SetTube("TUBE1", 
 "A992Fy50", 8, 6, 0.5, 0.5)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

This function is obsolete and has been superseded by
[SetTube\_1](../Definitions/Properties/Frame/SetTube_1.md)
as of version 24.2.

## See Also

[GetTube](GetTube_%7BFrame%7D.md)
