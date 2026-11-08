# GetPropertyFD

## Syntax

SapObject.SapModel.LinkObj.GetPropertyFD

## VB6 Procedure

Function GetProperty(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a defined link object.

PropName

The name of the frequency dependent link property assigned to the link object. This item is None if there is no frequency dependent link property assigned to the link object.

## Remarks

This function retrieves the frequency dependent link property assigned to a link object.

The function returns zero if the property is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkObjectFDProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropName As String



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



   'get link frequency dependent property

      ret = SapModel.LinkObj.GetPropertyFD("1", PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetProperty](SetProperty_%7BLink_Object%7D.md)

[GetProperty](GetProperty_%7BLink_Object%7D.md)

[SetPropertyFD](SetPropertyFD.md)
