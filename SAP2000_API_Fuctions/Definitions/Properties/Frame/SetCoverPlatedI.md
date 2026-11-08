# SetCoverPlatedI

## Syntax

SapObject.SapModel.PropFrame.SetCoverPlatedI

## VB6 Procedure

Function SetCoverPlatedI(ByVal Name As String, ByVal SectName As String, ByVal FyTopFlange As Double, ByVal FyWeb As Double, ByVal FyBotFlange As Double, ByVal tc As Double, ByVal bc As Double, ByVal MatPropTop As String, ByVal tcb As Double, ByVal bcb As Double, ByVal MatPropBot As String, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

SectName

The name of an existing I-type frame section property that is used for the I-section portion of the coverplated I section.

FyTopFlange

The yield strength of the top flange of the I-section. [F/L2]

If this item is 0, the yield strength of the I-section specified by the SectName item is used.

FyWeb

The yield strength of the web of the I-section. [F/L2]

If this item is 0, the yield strength of the I-section specified by the SectName item is used.

FyBotFlange

The yield strength of the bottom flange of the I-section. [F/L2]

If this item is 0, the yield strength of the I-section specified by the SectName item is used.

tc

The thickness of the top cover plate. [L]

If the tc or the bc item is less than or equal to 0, no top cover plate exists.

bc

The width of the top cover plate. [L]

If the tc or the bc item is less than or equal to 0, no top cover plate exists.

MatPropTop

The name of the material property for the top cover plate.

This item applies only if both the tc and the bc items are greater than 0.

tcb

The thickness of the bottom cover plate. [L]

If the tcb or the bcb item is less than or equal to 0, no bottom cover plate exists.

bcb

The width of the bottom cover plate. [L]

If the tcb or the bcb item is less than or equal to 0, no bottom cover plate exists.

MatPropBot

The name of the material property for the bottom cover plate.

This item applies only if both the tcb and the bcb items are greater than 0.

Color

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function initializes a cover plated I-type frame section property. If this function is called for an existing frame section property, all items for the section are reset to their default value.

The function returns zero if the section property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropCoverPlatedI()

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



   'set new I-type frame section property

      ret = SapModel.PropFrame.SetISection("ISEC", "A992Fy50", 24, 8, 0.5, 0.3, 8, 0.5)



   'set new cover plated I-type frame section property

      ret = SapModel.PropFrame.SetCoverPlatedI("CPI1", "ISEC", 0, 36, 0, 0.75, 14, "A992Fy50", 0.5, 6, "A992Fy50")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetCoverPlatedI](GetCoverPlatedI.md)
