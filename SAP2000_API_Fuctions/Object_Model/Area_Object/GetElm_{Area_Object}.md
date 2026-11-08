# GetElm

## Syntax

SapObject.SapModel.AreaObj.GetElm

## VB6 Procedure

Function GetElm(ByVal Name As String, ByRef nelm As Long, ByRef Elm() as String) As Long

## Parameters

Name

The name of an existing area object.

nelm

The number of area elements created from the specified area object.

Elm

An array that includes the name of a area element created from the specified area object.

## Remarks

This function retrieves the names of the area elements (analysis model area) associated with a specified area object in the object-based model.

This function returns zero if the area element information is successfully returned; otherwise it returns nonzero. An error occurs if the analysis model does not currently exist.

## VBA Example



```
Sub GetAreaElementInfoForAreaObject()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim nelm As Long

      Dim Elm() As String



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



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get area element information

      ret = SapModel.AreaObj.GetElm("1", nelm, Elm)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
