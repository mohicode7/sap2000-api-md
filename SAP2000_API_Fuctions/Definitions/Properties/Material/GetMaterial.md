# GetMaterial

## Syntax

SapObject.SapModel.PropMaterial.GetMaterial

## VB6 Procedure

Function GetMaterial(ByVal Name As String, ByRef MatType
As eMatType, ByRef Color As Long, ByRef Notes As String, ByRef GUID As
String) As Long

## Parameters

Name

The name of an existing material property.

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

The display color assigned to the material.

Notes

The notes, if any, assigned to the material.

GUID

The GUID (global unique identifier), if any, assigned
to the material.

## Remarks

This function retrieves some basic material property
data.

The function returns zero if the material is successfully
retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetBasicMatPropData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatType As eMatType

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Steel", 
 eMatType_Steel, -1, "API example test", "Default")



   'get basic material property data

      ret = SapModel.PropMaterial.GetMaterial("Steel", 
 MatType, Color, Notes, GUID)



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

[SetMaterial](SetMaterial.md)
