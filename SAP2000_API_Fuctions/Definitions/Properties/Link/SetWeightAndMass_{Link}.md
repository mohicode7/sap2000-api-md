# SetWeightAndMass

## Syntax

SapObject.SapModel.PropLink.SetWeightAndMass

## VB6 Procedure

Function SetWeightAndMass(ByVal Name As String, ByVal w As Double, ByVal m As Double, ByVal R1 As Double, ByVal R2 As Double, ByVal R3 As Double) As Long

## Parameters

Name

The name of an existing link property.

w

The weight of the link. [F]

m

The translational mass of the link. [M]

R1

The rotational inertia of the link about its local 1 axis. [ML2]

R2

The rotational inertia of the link about its local 2 axis. [ML2]

R3

The rotational inertia of the link about its local 3 axis. [ML2]

## Remarks

This function assigns weight and mass values to a link property.

The function returns zero if the values are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetLinkPropWeightAndMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double



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



   'add link property

      ReDim DOF(5)

      ReDim Fixed(5)

      ReDim Ke(5)

      ReDim Ce(5)

      DOF(0) = True

      Ke(0) = 12

      ret = SapModel.PropLink.SetLinear("L1", DOF, Fixed, Ke, Ce, 0, 0)



   'set link property weight and mass

      ret = SapModel.PropLink.SetWeightAndMass("L1", 10, 0.26, 0.0012, 0.0014, 0.0016)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetWeightAndMass](GetWeightAndMass_%7BLink%7D.md)
