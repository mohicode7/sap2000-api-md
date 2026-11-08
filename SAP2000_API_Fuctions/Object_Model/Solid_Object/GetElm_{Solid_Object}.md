# GetElm

## Syntax

SapObject.SapModel.SolidObj.GetElm

## VB6 Procedure

Function GetElm(ByVal Name As String, ByRef nelm As Long, ByRef Elm() as String) As Long

## Parameters

Name

The name of an existing solid object.

nelm

The number of solid elements created from the specified solid object.

Elm

An array that includes the name of a solid element created from the specified solid object.

## Remarks

This function retrieves the names of the solid elements (analysis model solid) associated with a specified solid object in the object-based model.

This function returns zero if the solid element information is successfully returned; otherwise it returns nonzero. An error occurs if the analysis model does not exist.

## VBA Example



```
Sub GetSolidElementInfoForSolidObject()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get solid element information

      ret = SapModel.SolidObj.GetElm("1", nelm, Elm)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
