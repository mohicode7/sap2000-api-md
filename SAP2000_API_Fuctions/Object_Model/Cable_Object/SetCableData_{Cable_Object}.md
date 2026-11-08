# SetCableData

## Syntax

SapObject.SapModel.CableObj.SetCableData

## VB6 Procedure

Function SetCableData(ByVal Name As String, ByVal CableType As Long, ByVal NumSegs As Long, ByVal Weight As Double, ByVal ProjectedLoad As Double, ByVal Value As Double, Optional ByVal UseDeformedGeom As Boolean = False, Optional ByVal ModelUsingFrames As Boolean = False) As Long

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

Value

This is the value of the parameter used to define the cable shape. The item that Value represents depends on the CableType item.

CableType = 1: Not Used

CableType = 2: Not Used

CableType = 3: Tension at I-End [F]

CableType = 4: Tension at J-End [F]

CableType = 5: Horizontal tension component [F]

CableType = 6: Maximum vertical sag [L]

CableType = 7: Low-point vertical sag [L]

CableType = 8: Undeformed length [L]

CableType = 9: Relative undeformed length

UseDeformedGeom

If this item is True, the program uses the deformed geometry for the cable object; otherwise it uses the undeformed geometry.

ModelUsingFrames

If this item is True, the analysis model uses frame elements to model the cable instead of using cable elements.

## Remarks

This function assigns the cable definition parameters to a cable object.

The function returns zero if the cable object is successfully defined; otherwise it returns a nonzero value. If the cable object is not successfully defined, it may be deleted.

## VBA Example



```
Sub SetCableDefinitionData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Modified optional arguments UseDeformedGeom and ModelUsingFrames to both be ByVal in version 12.0.1.

## See Also

[AddByCoord](AddByCoord_%7BCable_Object%7D.md)

[AddByPoint](AddByPoint_%7BCable_Object%7D.md)

[GetCableData](GetCableData_%7BCable_Object%7D.md)

[GetCableGeometry](GetCableGeometry_%7BCable_Object%7D.md)
