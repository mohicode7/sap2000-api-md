# ExtrudeFrameToAreaLinear

## Syntax

SapObject.SapModel.EditGeneral.ExtrudeFrameToAreaLinear

## VB6 Procedure

Function ExtrudeFrameToAreaLinear(ByVal Name As String,
ByVal PropName As String, ByVal dx As Double, ByVal dy As Double, ByVal
dz As Double, ByVal NumberAreas As Long, ByRef AreaName() As String, Optional
ByVal Remove As Boolean = True) As Long

## Parameters

Name

The name of an existing straight frame object to be
extruded.

PropName

This is Default, None or the name of a defined area
property to be used for the new extruded area objects.

dx, dy, dz

These are the x, y and z offsets used, in the present
coordinate system, to create each new area object.

NumberAreas

The number of area objects created when the specified
line object is extruded.

AreaName

This is an array of the name of each area object created
when the specified line object is extruded.

Remove

If this item is True, the straight frame object indicated
by the Name item is deleted after the extrusion is complete.

## Remarks

This function creates new area objects by linearly extruding
a specified straight frame object into area objects.

The function returns zero if the extrusion is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub LinearFrameExtrusionToAreas()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim AreaName() As String



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



   'linearly extrude frame to areas

      ret = SapModel.EditGeneral.ExtrudeFrameToAreaLinear("8", 
 "Default", 0, 144, 0, 3, AreaName, True)



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

[ExtrudeFrameToAreaRadial](ExtrudeFrameToAreaRadial.md)
