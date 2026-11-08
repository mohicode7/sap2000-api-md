# ExtrudePointToFrameLinear

## Syntax

SapObject.SapModel.EditGeneral.ExtrudePointToFrameLinear

## VB6 Procedure

Function ExtrudePointToFrameLinear(ByVal Name As String,
ByVal PropName As String, ByVal dx As Double, ByVal dy As Double, ByVal
dz As Double, ByVal NumberFrames As Long, ByRef FrameName() As String)
As Long

## Parameters

Name

The name of an existing point object to be extruded.

PropName

This is Default, None or the name of a defined frame
section property to be used for the new extruded frame objects.

dx, dy, dz

These are the x, y and z offsets used, in the present
coordinate system, to create each new frame object.

NumberFrames

The number of frame objects created when the specified
point object is extruded.

FrameName

This is an array of the name of each frame object created
when the specified point object is extruded.

## Remarks

This function creates new frame objects by linearly
extruding a specified point object into frame objects.

The function returns zero if the extrusion is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub LinearPointExtrusionToFrames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FrameName() As String



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



   'linearly extrude point to frames

      ret = SapModel.EditGeneral.ExtrudePointToFrameLinear("3", 
 "FSec1", 0, 144, 0, 3, FrameName)



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

## See Also

[ExtrudePointToFrameRadial](ExtrudePointToFrameRadial.md)
