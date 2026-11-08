# GetProp

## Syntax

SapObject.SapModel.PropTendon.GetProp

## VB6 Procedure

Function GetProp(ByVal Name As String, ByRef MatProp As String, ByRef ModelingOption As Long, ByRef Area As Double, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing tendon property.

MatProp

The name of the material property assigned to the tendon property.

ModelingOption

This is either 1 or 2, indicating the tendon modeling option.

1 = Model tendon as loads

2 = Model tendon as elements

Area

The cross-sectional area of the tendon. [L2]

Color

The display color assigned to the property.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property.

## Remarks

This function retrieves tendon property definition data.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonProperty()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MatProp As String

      Dim ModelingOption As Long

      Dim Area As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_TENDON, , , , , , MATERIAL_TENDON_SUBTYPE_ASTM_A416Gr270)



   'set new tendon property

      ret = SapModel.PropTendon.SetProp("T1", Name, 1, 2.25)



   'get tendon property data

      ret = SapModel.PropTendon.GetProp("T1", MatProp, ModelingOption, Area, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetProp](SetProp_%7BTendon%7D.md)
