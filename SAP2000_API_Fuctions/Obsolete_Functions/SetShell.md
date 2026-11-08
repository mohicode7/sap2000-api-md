# SetShell  (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropArea.SetShell

## VB6 Procedure

Function SetShell(ByVal Name As String, ByVal ShellType As Long, ByVal MatProp As String, ByVal MatAng As Double, ByVal Thickness As Double, ByVal Bending As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new area property. If this is an existing property, that property is modified; otherwise, a new property is added.

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

The display color assigned to the property. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, the program assigns a GUID to the property.

## Remarks

This function initializes a shell-type area property. If this function is called for an existing area property, all items for the property are reset to their default value.

The function returns zero if the property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAreaPropShell()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

The function is obsolete and has been superseded by [SetShell\_1](../Definitions/Properties/Area/SetShell_1.md) as of version 14.00. This function is maintained for backward compatibility. New function added

## See Also

[GetShell](GetShell.md)
