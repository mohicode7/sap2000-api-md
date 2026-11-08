# GetSectProps

## Syntax

SapObject.SapModel.PropFrame.GetSectProps

## VB6 Procedure

Function GetSectProps(ByVal Name As String, ByRef Area As Double, ByRef As2 As Double, ByRef As3 As Double, ByRef Torsion As Double, ByRef I22 As Double, ByRef I33 As Double, ByRef S22 As Double, ByRef S33 As Double, ByRef Z22 As Double, ByRef Z33 As Double, ByRef R22 As Double, ByRef R33 As Double) As Long

## Parameters

Name

The name of an existing frame section property.

Area

The cross-sectional area. [L2]

As2

The shear area for forces in the section local 2-axis direction. [L2]

As3

The shear area for forces in the section local 3-axis direction. [L2]

Torsion

The torsional constant. [L4]

I22

The moment of inertia for bending about the local 2 axis. [L4]

I33

The moment of inertia for bending about the local 3 axis. [L4]

S22

The section modulus for bending about the local 2 axis. [L3]

S33

The section modulus for bending about the local 3 axis. [L3]

Z22

The plastic modulus for bending about the local 2 axis. [L3]

Z33

The plastic modulus for bending about the local 3 axis. [L3]

R22

The radius of gyration about the local 2 axis. [L]

R33

The radius of gyration about the local 3 axis. [L]

## Remarks

This function retrieves properties for frame section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropProperties()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Area As Double

      Dim as2 As Double

      Dim as3 As Double

      Dim Torsion As Double

      Dim I22 As Double

      Dim I33 As Double

      Dim S22 As Double

      Dim S33 As Double

      Dim Z22 As Double

      Dim Z33 As Double

      Dim R22 As Double

      Dim R33 As Double



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



   'get frame section properties

      ret = SapModel.PropFrame.GetSectProps("FSEC1", Area, As2, As3, Torsion, I22, I33, S22, S33, Z22, Z33, R22, R33)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
