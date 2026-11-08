# GetSelected

## Syntax

Sap2000.AreaObj.GetSelected

## VB6 Procedure

Function GetSelected(ByVal Name As String, ByRef Selected As Boolean) As Long

## Parameters

Name

The name of an existing area object.

Selected

This item is True if the specified area object is selected; otherwise it is False.

## Remarks

This function retrieves the selected status for an area object.

The function returns zero if the selected status is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectSelectedStatus()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'set all area objects selected

      ret = SapModel.AreaObj.SetSelected("ALL", True, Group)



   'get area object selected status

      ret = SapModel.AreaObj.GetSelected("1", Selected)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSelected](SetSelected_%7BArea_Object%7D.md)

[GetSelectedEdge](GetSelectedEdge.md)

[SetSelectedEdge](SetSelectedEdge.md)
