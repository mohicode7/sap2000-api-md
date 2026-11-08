# ExtrudeAreaToSolidLinearNormal

## Syntax

SapObject.SapModel.EditGeneral.ExtrudeAreaToSolidLinearNormal

## VB6 Procedure

Function ExtrudeAreaToSolidLinearNormal(ByVal Name As
String, ByVal PropName As String, ByVal nPlus3 As Long, ByVal tPlus3 As
Double, ByVal nMinus3 As Long, ByVal tMinus3 As Double, ByRef NumberSolids
As Long, ByRef SolidName() As String, Optional ByVal Remove As Boolean
= True) As Long

## Parameters

Name

The name of an existing area object to be extruded.

PropName

This is either Default or the name of a defined solid
property to be used for the new extruded solid objects.

nPlus3

The number of solid objects created in the positive
local 3-axis direction of the specified area object.

tPlus3

The thickness of the solid objects created in the positive
local 3-axis direction of the specified area object.

nMinus3

The number of solid objects created in the negative
local 3-axis direction of the specified area object.

tMinus3

The thickness of the solid objects created in the negative
local 3-axis direction of the specified area object.

NumberSolids

The number of solid objects created when the specified
area object is extruded.

SolidName

This is an array of the name of each solid object created
when the specified area object is extruded.

Remove

If this item is True, the area object indicated by the
Name item is deleted after the extrusion is complete.

## Remarks

This function creates new solid objects by linearly
extruding a specified area object, in the local 3-axis direction of the
area object, into solid objects.

The function returns zero if the extrusion is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub LinearNormalAreaExtrusionToSolids()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberSolids As Long

      Dim SolidName() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(2, 48, 
 2, 48)



   'linearly extrude area to solids normal to area

      ret = SapModel.EditGeneral.ExtrudeAreaToSolidLinearNormal("2", 
 "Default", 1, 48, 1, 48, NumberSolids, SolidName, True)



   'refresh view

      ret = SapModel.View.RefreshView(0, 
 False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Modified optional argument Remove to be ByVal in version
12.0.1.

## See Also

[ExtrudeAreaToSolidLinearUser](ExtrudeAreaToSolidLinearUser.md)

[ExtrudeAreaToSolidRadial](ExtrudeAreaToSolidRadial.md)
