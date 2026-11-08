# GetCoordCylindrical

## Syntax

SapObject.SapModel.PointElm.GetCoordCylindrical

## VB6 Procedure

Function GetCoordCylindrical(ByVal Name As String, ByRef r As Double, ByRef Theta As Double, ByRef z As Double, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of an existing point element.

r

The radius for the point element in the specified coordinate system. [L]

Theta

The angle for the specified point element in the specified coordinate system. The angle is measured in the XY plane from the positive X axis. When looking in the XY plane with the positive Z axis pointing toward you, a positive Theta angle is counter clockwise. [deg]

z

The Z-coordinate of the specified point element in the specified coordinate system. [L]

CSys

The name of the coordinate system in which the joint coordinates are returned.

## Remarks

The function returns zero if the coordinates are successfully returned; otherwise it returns nonzero. If successful, the function returns the r, Theta and z coordinates of the specified point element in the Present Units. The coordinates are reported in the coordinate system specified by CSys.

## VBA Example



```
Sub GetPointElmCoordCylindrical()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim r As Double, Theta As Double, z As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get cylindrical point element coordinates

      ret = SapModel.PointElm.GetCoordCylindrical("5", r, Theta, z)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetCoordCartesian](GetCoordCartesian_%7BPoint_Element%7D.md)

[GetCoordSpherical](GetCoordSpherical_%7BPoint_Element%7D.md)
