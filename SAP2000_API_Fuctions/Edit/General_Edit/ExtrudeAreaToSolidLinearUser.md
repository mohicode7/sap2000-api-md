# ExtrudeAreaToSolidLinearUser

## Syntax

SapObject.SapModel.EditGeneral.ExtrudeAreaToSolidLinearUser

## VB6 Procedure

Function ExtrudeAreaToSolidLinearUser(ByVal Name As
String, ByVal PropName As String, ByVal dx As Double, ByVal dy As Double,
ByVal dz As Double, ByVal Number As Long, ByRef NumberSolids As Long,
ByRef SolidName() As String, Optional ByVal Remove As Boolean = True)
As Long

## Parameters

Name

The name of an existing area object to be extruded.

PropName

This is either Default or the name of a defined solid
property to be used for the new extruded solid objects.

dx, dy, dz

These are the x, y and z offsets used, in the present
coordinate system, to create each new solid object.

Number

The number of increments for the extrusion.

NumberSolids

The number of solid objects created when the specified
area object is extruded. Usually this item is returned the same as the
Number item. However, in some cases, such as when an area object with
more than four sides is extruded, this item will be larger than the Number
item.

SolidName

This is an array of the name of each solid object created
when the specified area object is extruded.

Remove

If this item is True, the area object indicated by the
Name item is deleted after the extrusion is complete.

## Remarks

This function creates new solid objects by linearly
extruding a specified area object, in a user specified direction, into
solid objects.

The function returns zero if the extrusion is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub LinearUserAreaExtrusionToSolids()

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



   'linearly extrude area to solids in user direction

      ret = SapModel.EditGeneral.ExtrudeAreaToSolidLinearUser("2", 
 "Default", 20, 144, 0, 3, NumberSolids, SolidName, True)



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

[ExtrudeAreaToSolidLinearNormal](ExtrudeAreaToSolidLinearNormal.md)

[ExtrudeAreaToSolidRadial](ExtrudeAreaToSolidRadial.md)
