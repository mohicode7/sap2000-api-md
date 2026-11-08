# GetTCLimits

## Syntax

SapObject.SapModel.LineElm.GetTCLimits

## VB6 Procedure

Function GetTCLimits(ByVal Name As String, ByRef LimitCompressionExists As Boolean, ByRef LimitCompression As Double, ByRef LimitTensionExists As Boolean, ByRef LimitTension As Double) As Long

## Parameters

Name

The name of an existing line element.

LimitCompressionExists

This item is True if a compression force limit exists for the line element.

LimitCompression

The compression force limit for the line element. [F]

LimitTensionExists

This item is True if a tension force limit exists for the line element.

LimitTension

The tension force limit for the line element. [F]

## Remarks

This function retrieves the tension/compression force limit assignments to line elements.

The function returns zero if the assignments are successfully retrieved, otherwise it returns a nonzero value.

Note that the tension and compression limits are only used in nonlinear analyses.

## VBA Example



```
Sub GetLineElmTCLimits()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim LimitCompressionExists As Boolean

      Dim LimitCompression As Double

      Dim LimitTensionExists As Boolean

      Dim LimitTension As Double



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



   'assign tension/compression limits to frame object

      ret = SapModel.FrameObj.SetTCLimits("1", True, -200, True, 30)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get tension/compression limits for line element

      ret = SapModel.LineElm.GetTCLimits("1-1", LimitCompressionExists, LimitCompression, LimitTensionExists, LimitTension)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
