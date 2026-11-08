# GetSelected

## Syntax

Sap2000.SolidObj.GetSelected

## VB6 Procedure

Function GetSelected(ByVal Name As String, ByRef Selected As Boolean) As Long

## Parameters

Name

The name of an existing solid object.

Selected

This item is True if the specified solid object is selected; otherwise it is False.

## Remarks

This function retrieves the selected status for a solid object.

The function returns zero if the selected status is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidObjectSelectedStatus()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'set all solid objects selected

      ret = SapModel.SolidObj.SetSelected("ALL", True, Group)



   'get solid object selected status

      ret = SapModel.SolidObj.GetSelected("1", Selected)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSelected](SetSelected_%7BSolid_Object%7D.md)
