# GetCoordCartesian

## Syntax

SapObject.SapModel.PointElm.GetCoordCartesian

## VB6 Procedure

Function GetCoordCartesian(ByVal Name As String, ByRef x As Double, ByRef y As Double, ByRef z As Double, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of an existing point element.

x

The X-coordinate of the specified point element in the specified coordinate system. [L]

y

The Y-coordinate of the specified point element in the specified coordinate system. [L]

z

The Z-coordinate of the specified point element in the specified coordinate system. [L]

CSys

The name of the coordinate system in which the joint coordinates are returned.

## Remarks

The function returns zero if the coordinates are successfully returned; otherwise it returns nonzero. If successful, the function returns the x, y and z coordinates of the specified point element in the Present Units. The coordinates are reported in the coordinate system specified by CSys.

## VBA Example



```
Sub GetPointElmCoordCartesian()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get cartesian point element coordinates

      ret = SapModel.PointElm.GetCoordCartesian("5", x, y, z)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetCoordCylindrical](GetCoordCylindrical_%7BPoint_Element%7Dhtm.md)

[GetCoordSpherical](GetCoordSpherical_%7BPoint_Element%7D.md)
