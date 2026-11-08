# GetDimensions

## Syntax

SapObject.SapModel.Options.GetDimensions

## VB6 Procedure

Function GetDimensions(ByRef CuttingPlaneTol As Double, ByRef WorldSpacing As Double, ByRef NudgeValue As Double, ByRef PixelClickSize As Long, ByRef PixelSnapSize As Long, ByRef ScrLinThk As Long, ByRef PrtLinThk As Long, ByRef MaxFont As Long, ByRef MinFont As Long, ByRef ZoomStep As Long, ByRef ShrinkFact As Long, ByRef TextFileMaxChar As Long) As Long

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

The auto zoom step size in percent (0 < ZoomStep < = 100).

ShrinkFact

The shrink factor in percent (0 < ShrinkFact < = 100).

TextFileMaxChar

The maximum line length in the text file (ShrinkFact > = 80).

## Remarks

This function retrieves the program dimension and tolerance items.

The function returns zero if the items are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetDimensionItems()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim CuttingPlaneTol As Double

      Dim WorldSpacing As Double

      Dim NudgeValue As Double

      Dim PixelClickSize As Long

      Dim PixelSnapSize As Long

      Dim ScrLinThk As Long

      Dim PrtLinThk As Long

      Dim MaxFont As Long

      Dim MinFont As Long

      Dim ZoomStep As Long

      Dim ShrinkFact As Long

      Dim TextFileMaxChar As Long



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



   'get dimensions and tolerances

      ret = SapModel.Options.GetDimensions(CuttingPlaneTol, WorldSpacing, NudgeValue, PixelClickSize, PixelSnapSize, ScrLinThk, PrtLinThk, MaxFont, MinFont, ZoomStep, ShrinkFact, TextFileMaxChar)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetDimensions](SetDimensions.md)
