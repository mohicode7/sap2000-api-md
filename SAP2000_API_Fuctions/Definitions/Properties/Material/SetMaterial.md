# SetMaterial

## Syntax

SapObject.SapModel.PropMaterial.SetMaterial

## VB6 Procedure

Function SetMaterial(ByVal Name As String, ByVal MatType
As eMatType, Optional ByVal Color As Long = -1, Optional ByVal Notes As
String = "", Optional ByVal GUID As String = "") As
Long

## Parameters

Name

The name of an existing or new material property. If
this is an existing property, that property is modified; otherwise, a
new property is added.

MatType

This is one of the following items in the eMatType enumeration.

eMatType\_Steel
= 1

eMatType\_Concrete
= 2

eMatType\_NoDesign
= 3

eMatType\_Aluminum
= 4

eMatType\_ColdFormed
= 5

eMatType\_Rebar
= 6

eMatType\_Tendon
= 7

Color

The display color assigned to the material. If Color
is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the material.

GUID

The GUID (global unique identifier), if any, assigned
to the material. If this item is input as Default, the program assigns
a GUID to the material.

## Remarks

This function initializes a material property. If this
function is called for an existing material property, all items for the
material are reset to their default value.

The function returns zero if the material is successfully
initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub InitializeMatProp()

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



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Steel", 
 eMatType_Steel)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Updated the documentation of the eMatType enumeration
in v22.1.0

Initial release in version 11.02.

## See Also

[GetMaterial](GetMaterial.md)
