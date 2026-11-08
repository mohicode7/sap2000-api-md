# GetShell\_1

## Syntax

SapObject.SapModel.PropArea.GetShell

## VB6 Procedure

Function GetShell\_1(ByVal Name As String, ByRef ShellType As Long, ByRef IncludeDrillingDOF As Boolean, ByRef MatProp As String, ByRef MatAng As Double, ByRef Thickness As Double, ByRef Bending As Double, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing shell-type area property.

ShellType

This is 1, 2, 3, 4, 5 or 6, indicating the shell type.

1 = Shell - thin

2 = Shell - thick

3 = Plate - thin

4 = Plate - thick

5 = Membrane

6 = Shell layered/nonlinear

IncludeDrillingDOF

This item is True if drilling degrees of freedom are included in the element formulation in the analysis model. This item does not apply when ShellType = 3, 4 or 6.

MatProp

The name of the material property for the area property. This item does not apply when ShellType = 6.

MatAng

The material angle. [deg]

This item does not apply when ShellType = 6.

Thickness

The membrane thickness. [L]

This item does not apply when ShellType = 6.

Bending

The bending thickness. [L]

This item does not apply when ShellType = 6.

Color

The display color assigned to the property.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property.

## Remarks

This function retrieves area property data for a shell-type area section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaPropShell_1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ShellType As Long

      Dim IncludeDrillingDOF As Boolean

      Dim MatProp As String

      Dim MatAng As Double

      Dim Thickness As Double

      Dim Bending As Double

      Dim Color As Long

      Dim Notes As String

      Dim GUID As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'set new area property

      ret = SapModel.PropArea.SetShell("A1", 1, "4000Psi", 0, 16, 16)



   'get area property data

      ret = SapModel.PropArea.GetShell_1("A1", ShellType, IncludeDrillingDOF, MatProp, MatAng, Thickness, Bending, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

This function supersedes [GetShell](../../../Obsolete_Functions/GetShell.md).

## See Also

[SetShell\_1](SetShell_1.md)
