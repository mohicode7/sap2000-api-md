# GetShell (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropArea.GetShell

## VB6 Procedure

Function GetShell(ByVal Name As String, ByRef ShellType As Long, ByRef MatProp As String, ByRef MatAng As Double, ByRef Thickness As Double, ByRef Bending As Double, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

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
Sub GetAreaPropShell()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ShellType As Long

      Dim MatProp As String

      Dim MatAng As Double

      Dim Thickness As Double

      Dim Bending As Double

      Dim Color As Long

      Dim Notes As String

      Dim GUID As String



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



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

      ret = SapModel.PropArea.GetShell("A1", ShellType, MatProp, MatAng, Thickness, Bending, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

The function is obsolete and has been superseded by [GetShell\_1](../Definitions/Properties/Area/GetShell_1.md) as of version 14.00. This function is maintained for backward compatibility. New function added

## See Also

[SetShell](SetShell.md)
