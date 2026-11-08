# SetShell\_1

## Syntax

SapObject.SapModel.PropArea.SetShell\_1

## VB6 Procedure

Function SetShell(ByVal Name As String, ByVal ShellType As Long, ByVal IncludeDrillingDOF As Boolean, ByVal MatProp As String, ByVal MatAng As Double, ByVal Thickness As Double, ByVal Bending As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

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

The display color assigned to the property. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, then the program assigns a GUID to the property.

## Remarks

This function initializes a shell-type area property. If this function is called for an existing area property, all items for the property are reset to their default value.

The function returns zero if the property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAreaPropShell_1()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'set new area property

      ret = SapModel.PropArea.SetShell_1("A1", 1, True, "4000Psi", 0, 16, 16)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

This function supersedes [SetShell](../../../Obsolete_Functions/SetShell.md).

## See Also

[GetShell\_1](GetShell_1.md)
