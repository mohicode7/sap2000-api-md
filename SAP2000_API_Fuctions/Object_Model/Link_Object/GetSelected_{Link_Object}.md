# GetSelected

## Syntax

Sap2000.LinkObj.GetSelected

## VB6 Procedure

Function GetSelected(ByVal Name As String, ByRef Selected As Boolean) As Long

## Parameters

Name

The name of an existing link object.

Selected

This item is True if the specified link object is selected; otherwise it is False.

## Remarks

This function retrieves the selected status for a link object.

The function returns zero if the selected status is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkObjectSelectedStatus()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim Selected As Boolean



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



   'set link object selected

      ret = SapModel.LinkObj.SetSelected("ALL", True, Group)



   'get link object selected status

      ret = SapModel.LinkObj.GetSelected(Name, Selected)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSelected](SetSelected_%7BLink_Object%7D.md)
