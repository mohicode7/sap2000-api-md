# SetCoordSys

## Syntax

SapObject.SapModel.CoordSys.SetCoordSys

## VB6 Procedure

Function SetCoordSys(ByVal Name As String, ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal RZ As Double, ByVal RY As Double, ByVal RX As Double) As Long

## Parameters

Name

This is the name of a coordinate system. If this is the name of an existing coordinate system, that coordinate system is modified, otherwise a new coordinate system is added.

x

The global X coordinate of the origin of the coordinate system. [L]

y

The global Y coordinate of the origin of the coordinate system. [L]

z

The global Z coordinate of the origin of the coordinate system. [L]

RZ, RY, RX

The rotation of an axis of the new coordinate system relative to the global coordinate system is defined as follows: (1) Rotate the coordinate system about the positive global Z-axis as defined by the RZ item. (2) Rotate the coordinate system about the positive global Y-axis as defined by the RY item. (3) Rotate the coordinate system about the positive global X-axis as defined by the RX item.  Note that the order in which these rotations are performed is important.  RX, RY and RZ are angles in degrees [deg].

## Remarks

The function returns zero if the coordinate system is successfully added or modified, otherwise it returns a nonzero value.

Modifying the Global coordinate system will fail and return an error.

## VBA Example



```
Sub AddCoordSystem()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetCoordSys](GetCoordSys.md)

[Delete](Delete_%7BCoordinate_System%7D.md)
