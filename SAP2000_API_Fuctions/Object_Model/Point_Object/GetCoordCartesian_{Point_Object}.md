# GetCoordCartesian

## Syntax

SapObject.SapModel.PointObj.GetCoordCartesian

## VB6 Procedure

Function GetCoordCartesian(ByVal Name As String, ByRef x As Double, ByRef y As Double, ByRef z As Double, Optional ByVal Csys As String = "Global") As Long

## Parameters

Name

The name of a defined point object.

x

The X-coordinate of the specified point object in the specified coordinate system. [L]

y

The Y-coordinate of the specified point object in the specified coordinate system. [L]

z

The Z-coordinate of the specified point object in the specified coordinate system. [L]

Csys

The name of a defined coordinate system. If Csys is not specified, the Global coordinate system is assumed.

## Remarks

The function returns zero if the coordinates are successfully returned; otherwise it returns nonzero. If successful, the function returns the x, y and z coordinates of the specified point object in the Present Units. The coordinates are reported in the coordinate system specified by Csys.

## VBA Example



```
Sub GetPointCoordCartesian

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim x As Double, y As Double, z As Double

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

      ret = SapModel.PointObj.GetCoordCartesian("1", x, y, z)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetCoordCylindrical](GetCoordCylindrical_%7BPoint_Object%7D.md)

[GetCoordSpherical](GetCoordSpherical_%7BPoint_Object%7D.md)

[GetNameList](GetNameList_%7BPoint_Object%7D.md)
