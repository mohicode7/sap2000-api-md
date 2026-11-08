# GetCoordCylindrical

## Syntax

SapObject.SapModel.PointObj.GetCoordCylindrical

## VB6 Procedure

Function GetCoordCylindrical(ByVal Aame As String, ByRef r As Double, ByRef Theta As Double, ByRef z As Double, Optional ByVal Csys As String = "Global") As Long

## Parameters

Name

The name of a defined point object.

r

The radius for the specified point object in the specified coordinate system. [L]

Theta

The angle for the specified point object in the specified coordinate system. The angle is measured in the XY plane from the positive X axis. When looking in the XY plane with the positive Z axis pointing toward you, a positive Theta angle is counter clockwise [deg]

z

The Z-coordinate of the specified point object in the specified coordinate system. [L]

Csys

The name of a defined coordinate system. If Csys is not specified, the Global coordinate system is assumed.

## Remarks

The function returns zero if the coordinates are successfully returned; otherwise it returns nonzero. If successful, the function returns the r, theta and z coordinates of the specified point object in the Present Units. The coordinates are reported in the coordinate system specified by CSys.

## VBA Example



```
Sub GetPointCoordCylindrical

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



   'start a new template model

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'get point coordinates of the point named "1" and display in a message box

      ret = SapModel.PointObj.GetCoordCylindrical("1", r, Theta, z)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetCoordCartesian](GetCoordCartesian_%7BPoint_Object%7D.md)

[GetCoordSpherical](GetCoordSpherical_%7BPoint_Object%7D.md)

[GetNameList](GetNameList_%7BPoint_Object%7D.md)
