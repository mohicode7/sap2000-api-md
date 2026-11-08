# GetCableGeometry

## Syntax

SapObject.SapModel.CableObj.GetCableGeometry

## VB6 Procedure

Function GetCableGeometry(ByRef Name As String, ByRef NumberPoints As Long, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, ByRef Sag() As Double, ByRef Dist() As Double, ByRef RD() As Double, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of a defined cable object.

NumberPoints

The number of points defining the cable geometry.

x, y, z

The x, y and z coordinates of the considered point on the cable in the coordinate system specified by the CSys item. [L]

Sag

The cable vertical sag, measured from the chord, at the considered point. [L]

Distance

The distance along the cable, measured from the cable I-End, to the considered point. [L]

RD

The relative distance along the cable, measured from the cable I-End, to the considered point.

CSys

The name of the coordinate system in which the x, y and z coordinates are to be reported.

## Remarks

This function retrieves geometric data for a specified cable object.

The function returns zero if the data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableObjectGeometry()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberPoints As Long

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double

      Dim Sag() As Double

      Dim Dist() As Double

      Dim RD() As Double



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



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'get cable geometry

      ret = SapModel.CableObj.GetCableGeometry(Name, NumberPoints , x, y, z, Sag, Dist, RD)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByCoord](AddByCoord_%7BCable_Object%7D.md)

[AddByPoint](AddByPoint_%7BCable_Object%7D.md)

[SetCableData](SetCableData_%7BCable_Object%7D.md)

[GetCableData](GetCableData_%7BCable_Object%7D.md)
