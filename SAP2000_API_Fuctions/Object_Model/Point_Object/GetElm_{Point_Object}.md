# GetElm

## Syntax

SapObject.SapModel.PointObj.GetElm

## VB6 Procedure

Function GetElm(ByVal Name As String, ByRef Elm as String) As Long

## Parameters

Name

The name of an existing point object.

Elm

The name of the point element associated with the specified point object.

## Remarks

This function retrieves the name of the point element (analysis model point) associated with a specified point object in the object-based model.

This function returns zero if the point element name is successfully returned; otherwise it returns nonzero. An error occurs if the analysis model does not currently exist.

## VBA Example



```
Sub GetPointElementName()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Elm As String



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



   'get point element name

      ret = SapModel.PointObj.GetElm("5", Elm)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
