# GetCoverPlatedI

## Syntax

SapObject.SapModel.PropFrame.GetCoverPlatedI

## VB6 Procedure

Function GetCoverPlatedI(ByVal Name As String, ByRef SectName As String, ByRef FyTopFlange As Double, ByRef FyWeb As Double, ByRef FyBotFlange As Double, ByRef tc As Double, ByRef bc As Double, ByRef MatPropTop As String, ByRef tcb As Double, ByRef bcb As Double, ByRef MatPropBot As String, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing frame section property.

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

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for a tube-type frame section.

The function returns zero if the section property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropCoverPlatedI()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SectName As String

      Dim FyTopFlange As Double

      Dim FyWeb As Double

      Dim FyBotFlange As Double

      Dim tc As Double

      Dim bc As Double

      Dim MatPropTop As String

      Dim tcb As Double

      Dim bcb As Double

      Dim MatPropBot As String

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set new I-type frame section property

      ret = SapModel.PropFrame.SetISection("ISEC", "A992Fy50", 24, 8, 0.5, 0.3, 8, 0.5)



   'set new cover plated I-type frame section property

      ret = SapModel.PropFrame.SetCoverPlatedI("CPI1", "ISEC", 0, 36, 0, 0.75, 14, "A992Fy50", 0.5, 6, "A992Fy50")



   'get frame section property data for cover plated I

      ret = SapModel.PropFrame.GetCoverPlatedI("CPI1", SectName, FyTopFlange, FyWeb, FyBotFlange, tc, bc, MatPropTop, tcb, bcb, MatPropBot, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetCoverPlatedI](SetCoverPlatedI.md)
