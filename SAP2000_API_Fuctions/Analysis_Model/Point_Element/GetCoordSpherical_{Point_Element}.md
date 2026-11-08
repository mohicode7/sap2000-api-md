# GetCoordSpherical

## Syntax

SapObject.SapModel.PointElm.GetCoordSpherical

## VB6 Procedure

Function GetCoordSpherical(ByVal Name As String, ByRef r As Double, ByRef a As Double, ByRef b As Double, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of an existing point element.

r

The radius for the point element in the specified coordinate system. [L]

a

The plan angle for the point element in the specified coordinate system. This angle is measured in the XY plane from the positive global X axis. When looking in the XY plane with the positive Z axis pointing toward you, a positive a angle is counter clockwise. [deg]

b

The elevation angle for the point element in the specified coordinate system. This angle is measured in an X'Z plane that is perpendicular to the XY plane with the positive X' axis oriented at angle a from the positive global X axis. Angle b is measured from the positive global Z axis. When looking in the X’Z plane with the positive Y' axis pointing toward you, a positive b angle is counter clockwise. [deg]

CSys

The name of the coordinate system in which the joint coordinates are returned.

## Remarks

The function returns zero if the coordinates are successfully returned; otherwise it returns nonzero. If successful, the function returns the r, a and b coordinates of the specified point element in the Present Units. The coordinates are reported in the coordinate system specified by CSys.

## VBA Example



```
Sub GetPointElmCoordSpherical()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim r As Double, a As Double, b As Double

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



   'get spherical point element coordinates

      ret = SapModel.PointElm.GetCoordSpherical("5", r, a, b)



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

[GetCoordCylindrical](GetCoordCylindrical_%7BPoint_Element%7Dhtm.md)
