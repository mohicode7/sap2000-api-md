# GetWeightAndMass

## Syntax

SapObject.SapModel.PropLink.GetWeightAndMass

## VB6 Procedure

Function GetWeightAndMass(ByVal Name As String, ByRef w As Double, ByRef m As Double, ByRef R1 As Double, ByRef R2 As Double, ByRef R3 As Double) As Long

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

This function retrieves weight and mass data for a link property.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkPropWeightAndMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

      Dim w As Double

      Dim m As Double

      Dim R1 As Double

      Dim R2 As Double

      Dim R3 As Double



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



   'get link property weight and mass

      ret = SapModel.PropLink.GetWeightAndMass("L1", w, m, R1, R2, R3)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetWeightAndMass](SetWeightAndMass_%7BLink%7D.md)
