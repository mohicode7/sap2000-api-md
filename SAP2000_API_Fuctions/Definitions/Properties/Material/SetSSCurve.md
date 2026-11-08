# SetSSCurve

## Syntax

SapObject.SapModel.PropMaterial.SetSSCurve

## VB6 Procedure

Function SetSSCurve(ByVal Name As String, ByVal NumberPoints As Long, ByRef PointID() As Long, ByRef Strain() As Double, ByRef Stress() As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing material property.

NumberPoints

The number of points in the stress-strain curve. This item must be at least 3.

PointID

This is one of the following integers which sets the point ID. The point ID controls the color that will be displayed for hinges in a deformed shape plot.

-5 = -E

-4 = -D

-3 = -C

-2 = -B

0 = None

1 = A

2 = B

3 = C

4 = D

5 = E

The point IDs must be input in numerically increasing order, except that 0 (None) values are allowed anywhere. No duplicate values are allowed excepth for 0 (None).

Strain

This is an array that includes the strain at each point on the stress strain curve. The strains must increase monotonically.

Stress

This is an array that includes the stress at each point on the stress strain curve. [F/L2]

Points that have a negative strain must have a zero or negative stress. Similarly, points that have a positive strain must have a zero or positive stress.

There must be one point defined that has zero strain and zero stress.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

## Remarks

This function sets the material stress-strain curve for materials that are specified to have user-defined stress-strain curves.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropUserSS()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PointID() As Long

      Dim Strain() As Double

      Dim Stress() As Double



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



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Steel", MATERIAL_STEEL)



   'assign other properties

      ret = SapModel.PropMaterial.SetOSteel("Steel", 55, 68, 60, 70, 0, 1, 0, 0, 0)



   'assign user SS curve

      ReDim PointID(4)

      ReDim Strain(4)

      ReDim Stress(4)

      Strain(0) = -0.003:  Stress(0) = -50:  PointID(0) = -3

      Strain(1) = -0.001:  Stress(1) = -25:  PointID(1) = 0

      Strain(2) = 0:       Stress(2) = -0:   PointID(2) = 1

      Strain(3) = 0.003:   Stress(3) = 40:   PointID(3) = 0

      Strain(4) = 0.008:   Stress(4) = 80:   PointID(4) = 5

      ret = SapModel.PropMaterial.SetSSCurve("Steel", 5, PointID, Strain, Stress)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetSSCurve](GetSSCurve.md)
