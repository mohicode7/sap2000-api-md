# GetCoordSys

## Syntax

SapObject.SapModel.CoordSys.GetCoordSys

## VB6 Procedure

Function GetCoordSys(ByVal Name As String, ByRef x As Double, ByRef y As Double, ByRef z As Double, ByRef RZ As Double, ByRef RY As Double, ByRef RX As Double) As Long

## Parameters

Name

The name of an existing coordinate system.

x

The global X coordinate of the origin of the coordinate system. [L]

y

The global Y coordinate of the origin of the coordinate system. [L]

z

The global Z coordinate of the origin of the coordinate system. [L]

RZ, RY, RX

The rotation of an axis of the new coordinate system relative to the global coordinate system is defined as follows: (1) Rotate the coordinate system about the positive global Z-axis as defined by the RZ item. (2) Rotate the coordinate system about the positive global Y-axis as defined by the RY item. (3) Rotate the coordinate system about the positive global X-axis as defined by the RX item.  Note that the order in which these rotations are performed is important.  RX, RY and RZ are angles in degrees [deg].

## Remarks

The function returns zero if the coordinate system data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCoordSystem()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim x as Double, y as Double, z as Double

      Dim RX as Double, RY as Double, RZ as Double



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



   'define new coordinate system

      ret = SapModel.CoordSys.SetCoordSys("CSys1", 1000, 1000, 0, 0, 0, 0)



   'get new coordinate system data

      ret = SapModel.CoordSys.GetCoordSys("CSys1", x, y, z, RZ, RY, RX)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetCoordSys](SetCoordSys.md)
