# GetPropertyFD

## Syntax

SapObject.SapModel.LinkElm.GetPropertyFD

## VB6 Procedure

Function GetPropertyFD(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of an existing link element.

PropName

The name of the frequency dependent link property assigned to the link element.

## Remarks

This function retrieves the frequency dependent property assignment to a link element. If no frequency dependent property is assigned to the link, the PropName is returned as None.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkElementFDProp()

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



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 6-012.sdb")



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get frequency dependent property for link element

      ret = SapModel.LinkElm.GetPropertyFD("1", PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
