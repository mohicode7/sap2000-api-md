# GetElm

## Syntax

SapObject.SapModel.TendonObj.GetElm

## VB6 Procedure

Function GetElm(ByVal Name As String, ByRef nelm As Long, ByRef Elm() as String, ByRef RDI() As Double, ByRef RDJ() As Double) As Long

## Parameters

Name

The name of an existing tendon object.

nelm

The number of line elements created from the specified tendon object.

Elm

An array that includes the name of a line element created from the specified tendon object.

RDI

An array that includes the relative distance along the tendon object to the I-End of the line element.

RDJ

An array that includes the relative distance along the tendon object to the J-End of the line element.

## Remarks

This function retrieves the names of the line elements (analysis model lines) associated with a specified tendon object in the object-based model. It also retrieves information about the location of the line elements along the tendon object.

This function returns zero if the line element information is successfully returned; otherwise it returns nonzero. An error occurs if the analysis model does not currently exist.

## VBA Example



```
Sub GetLineElementInfoForTendonObject()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim nelm As Long

      Dim Elm() As String

      Dim RDI() As Double

      Dim RDJ() As Double

      Dim Name As String



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



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'assign tendon force load

      ret = SapModel.TendonObj.SetLoadForceStress("ALL", "DEAD", 1, 0, 100, 0.15, 8.333E-05, 0.25, 3, 5, 7, 5, , Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element information

      ret = SapModel.TendonObj.GetElm(Name, nelm, Elm, RDI, RDJ)



   'Note:  The above returns 0 elements because the tendon is specified to be modelled using loads, not elements



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
