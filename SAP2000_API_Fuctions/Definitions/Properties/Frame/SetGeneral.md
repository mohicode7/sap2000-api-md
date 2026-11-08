# SetGeneral

## Syntax

SapObject.SapModel.PropFrame.SetGeneral

## VB6 Procedure

Function SetGeneral(ByVal Name As String, ByVal MatProp As String, ByVal t3 As Double, ByVal t2 As Double, ByVal Area As Double, ByVal As2 As Double, ByVal As3 As Double, ByVal Torsion As Double, ByVal I22 As Double, ByVal I33 As Double, ByVal S22 As Double, ByVal S33 As Double, ByVal Z22 As Double, ByVal Z33 As Double, ByVal R22 As Double, ByVal R33 As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

MatProp

The name of the material property for the section.

t3

The section depth. [L]

t2

The section width. [L]

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

Color

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function initializes a general frame section property. If this function is called for an existing frame section property, all items for the section are reset to their default value.

The function returns zero if the section property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropGeneral()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'set new frame section property

      ret = SapModel.PropFrame.SetGeneral("GEN1", "A992Fy50", 24, 14, 100, 80, 80, 4, 1000, 2400, 140, 200, 150, 220, 3, 5, -1, "API example", "Default")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetGeneral](GetGeneral.md)
