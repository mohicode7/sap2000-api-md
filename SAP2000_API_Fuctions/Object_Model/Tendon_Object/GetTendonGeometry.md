# GetTendonGeometry

## Syntax

SapObject.SapModel.TendonObj.GetTendonGeometry

## VB6 Procedure

Function GetTendonGeometry(ByVal Name As String, ByRef NumberPoints As Long, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of a defined tendon object.

NumberPoints

The number of items used to define the discretized tendon geometry.

x

This is an array of the X (or local 1) coordinate of each point in the coordinate system specified by CSys. [L]

y

This is an array of the Y (or local 2) coordinate of each point in the coordinate system specified by CSys. [L]

z

This is an array of the Z (or local 3) coordinate of each point in the coordinate system specified by CSys. [L]

CSys

This is the coordinate system in which the x, y and z coordinate parameters are defined. It is Local or the name of a defined coordinate system.

Local means that the point coordinates are in the local system of the specified tendon object with the origin assumed to be at the I-End of the tendon.

## Remarks

This function retrieves tendon object discretized geometry, that is, it retrieves the coordinates of the points along the discretized tendon.

The function returns zero if the tendon object discretized geometry is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonDiscretizedGeometry()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberPoints As Long

      Dim MyType() As Long

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double



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



   'add tendon object by coordinates

      ret = SapModel.TendonObj.AddByCoord(-288, 0, 288, 288, 0, 288, Name)



   'set tendon geometry

      NumberPoints = 3

      ReDim MyType(NumberPoints – 1)

      ReDim x(NumberPoints – 1)

      ReDim y(NumberPoints – 1)

      ReDim z(NumberPoints – 1)

      MyType(0) = 1

      MyType(1) = 7

      MyType(2) = 6

      x(0) = 0:  y(0) = 0

      x(1) = 288:  y(1) = -12

      x(2) = 576:  y(2) = 0

      ret = SapModel.TendonObj.SetTendonData(Name, NumberPoints, MyType, x, y, z, "Local")



   'update view

      ret = SapModel.View.RefreshView(0, False)



   'get tendon discretized geometry

      ReDim x(0)

      ReDim y(0)

      ReDim z(0)

      ret = SapModel.TendonObj.GetTendonGeometry(Name, NumberPoints, x, y, z, "Global")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetTendonData](GetTendonData.md)

[SetTendonData](SetTendonData.md)
