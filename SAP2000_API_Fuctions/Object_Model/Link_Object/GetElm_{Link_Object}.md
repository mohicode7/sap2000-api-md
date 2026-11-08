# GetElm

## Syntax

SapObject.SapModel.LinkObj.GetElm

## VB6 Procedure

Function GetElm(ByVal Name As String, ByRef Elm as String) As Long

## Parameters

Name

The name of an existing link object.

Elm

The name of the link element created from the specified link object.

## Remarks

This function retrieves the name of the link element (analysis model link) associated with a specified link object in the object-based model.

This function returns zero if the link element name is successfully retrieved; otherwise it returns nonzero. An error occurs if the analysis model does not exist.

## VBA Example



```
Sub GetLinkElementNameForLinkObject()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get link element name

      ret = SapModel.LinkObj.GetElm(Name, Elm)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
