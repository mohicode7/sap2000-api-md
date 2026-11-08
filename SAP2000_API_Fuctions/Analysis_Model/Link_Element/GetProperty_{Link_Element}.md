# GetProperty

## Syntax

SapObject.SapModel.LinkElm.GetProperty

## VB6 Procedure

Function GetProperty(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of an existing link element.

PropName

The name of the link property assigned to the link element.

## Remarks

This function retrieves the property assignment to a link element.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

The sVarTotalLength and sVarRelStartLoc items apply only when the Var item is True.

## VBA Example



```
Sub GetLinkElementProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropName As String

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



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name)



   'refresh view

      ret = SapModel.View.RefreshView



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get property for link element

      ret = SapModel.LinkElm.GetProperty(Name, PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
