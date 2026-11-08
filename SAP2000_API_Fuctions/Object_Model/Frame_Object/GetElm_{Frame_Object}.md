# GetElm

## Syntax

SapObject.SapModel.FrameObj.GetElm

## VB6 Procedure

Function GetElm(ByVal Name As String, ByRef nelm As Long, ByRef Elm() as String, ByRef RDI() As Double, ByRef RDJ() As Double) As Long

## Parameters

Name

The name of an existing frame object.

nelm

The number of line elements created from the specified frame object.

Elm

An array that includes the name of a frame element created from the specified frame object.

RDI

An array that includes the relative distance along the frame object to the I-End of the frame element.

RDJ

An array that includes the relative distance along the line object to the J-End of the frame element.

## Remarks

This function retrieves the names of the frame elements (analysis model lines) associated with a specified frame object in the object-based model. It also retrieves information about the location of the frame elements along the frame object.

This function returns zero if the frame element information is successfully returned; otherwise it returns nonzero. An error occurs if the analysis model does not exist.

## VBA Example



```
Sub GetLineElementInfo()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim nelm As Long

      Dim Elm() As String

      Dim RDI() As Double

      Dim RDJ() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element information

      ret = SapModel.FrameObj.GetElm("5", nelm, Elm, RDI, RDJ)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
