# GetCableData

## Syntax

SapObject.SapModel.CableObj.GetCableData

## VB6 Procedure

Function GetCableData(ByRef Name As String, ByRef CableType As Long, ByRef NumSegs As Long, ByRef Weight As Double, ByRef ProjectedLoad As Double, ByRef UseDeformedGeom As Boolean, ByRef ModelUsingFrames As Boolean, ByRef Parameter() As Double) As Long

## Parameters

Name

The name of a defined cable object.

CableType

This is 1, 2, 3, 4, 5, 6, 7, 8, or 9, indicating the cable definition parameter.

1 = Minimum tension at I-End

2 = Minimum tension at J-End

3 = Tension at I-End

4 = Tension at J-End

5 = Horizontal tension component

6 = Maximum vertical sag

7 = Low-point vertical sag

8 = Undeformed length

9 = Relative undeformed length

NumSegs

This is the number of segments into which the program internally divides the cable.

Weight

The added weight per unit length used when calculating the cable shape. [F/L]

ProjectedLoad

The projected uniform gravity load used when calculating the cable shape. [F/L]

UseDeformedGeom

If this item is True, the program uses the deformed geometry for the cable object; otherwise it uses the undeformed geometry.

ModelUsingFrames

If this item is True, the analysis model uses frame elements to model the cable instead of using cable elements.

Parameter

This is an array of parameters related to the cable shape. The array is dimensioned by Sap2000.

Parameter(0) = Tension at I-End [F]

Parameter(1) = Tension at J-End [F]

Parameter(2) = Horizontal tension component [F]

Parameter(3) = Maximum deformed vertical sag [L]

Parameter(4) = Deformed low-point vertical sag [L]

Parameter(5) = Deformed length [L]

Parameter(6) = Deformed relative length

Parameter(7) = Maximum undeformed vertical sag [L]

Parameter(8) = Undeformed low-point vertical sag [L]

Parameter(9) = Undeformed length [L]

Parameter(10) = Undeformed relative length

## Remarks

This function retrieves definition data for a specified cable object.

The function returns zero if the data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableObjectData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim CableType As Long

      Dim NumSegs As Long

      Dim Weight As Double

      Dim ProjectedLoad As Double

      Dim UseDeformedGeom As Boolean

      Dim ModelUsingFrames As Boolean

      Dim Parameter() As Double



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



   'get cable data

      ret = SapModel.CableObj.GetCableData(Name, CableType, NumSegs, Weight, ProjectedLoad, UseDeformedGeom, ModelUsingFrames, Parameter)



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

[GetCableGeometry](GetCableGeometry_%7BCable_Object%7D.md)
