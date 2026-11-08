# SetDimensions

## Syntax

SapObject.SapModel.Options.SetDimensions

## VB6 Procedure

Function SetDimensions(Optional ByVal CuttingPlaneTol As Double = 0, Optional ByVal WorldSpacing As Double = 0, Optional ByVal NudgeValue As Double = 0, Optional ByVal PixelClickSize As Long = 0, Optional ByVal PixelSnapSize As Long = 0, Optional ByVal ScrLinThk As Long = 0, Optional ByVal PrtLinThk As Long = 0, Optional ByVal MaxFont As Long = 0, Optional ByVal MinFont As Long = 0, Optional ByVal ZoomStep As Long = 0, Optional ByVal ShrinkFact As Long = 0, Optional ByVal TextFileMaxChar As Long = 0) As Long

## Parameters

CuttingPlaneTol

The tolerance for 2D view cutting planes. [L]

WorldSpacing

The plan fine grid spacing. [L]

NudgeValue

The plan nudge value. [L]

PixelClickSize

The screen selection tolerance in pixels.

PixelSnapSize

The screen snap tolerance in pixels.

ScrLinThk

The screen line thickness in pixels.

PrtLinThk

The printer line thickness in pixels.

MaxFont

The maximum graphic font size in points.

MinFont

The minimum graphic font size in points.

ZoomStep

The auto zoom step size in percent (0 < ZoomStep <= 100).

ShrinkFact

The shrink factor in percent (0 < ShrinkFact <= 100).

TextFileMaxChar

The maximum line length in the text file (ShrinkFact >= 80).

## Remarks

This function sets program dimension and tolerance items. Inputting 0 for any item means that the item will be ignored by the program; that is, its current value will not be changed.

The function returns zero if the items are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetDimensionItems()

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



   'set dimensions and tolerances

      ret = SapModel.Options.SetDimensions(12, 2, 1, 4, , , , , , , , 120)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetDimensions](GetDimensions.md)
