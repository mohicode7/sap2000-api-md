# GetElm

## Syntax

SapObject.SapModel.CableObj.GetElm

## VB6 Procedure

Function GetElm(ByVal Name As String, ByRef nelm As Long, ByRef Elm() as String, ByRef RDI() As Double, ByRef RDJ() As Double) As Long

## Parameters

Name

The name of an existing cable object.

nelm

The number of line elements created from the specified cable object.

Elm

An array that includes the name of a line element created from the specified cable object.

RDI

An array that includes the relative distance along the cable object to the I-End of the line element.

RDJ

An array that includes the relative distance along the cable object to the J-End of the line element.

## Remarks

This function retrieves the names of the line elements (analysis model lines) associated with a specified cable object in the object-based model. It also retrieves information about the location of the line elements along the cable object.

This function returns zero if the line element information is successfully returned; otherwise it returns nonzero. An error occurs if the analysis model does not currently exist.

## VBA Example



```
Sub GetLineElementInfoForCableObject()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element information

      ret = SapModel.CableObj.GetElm(Name, nelm, Elm, RDI, RDJ)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
