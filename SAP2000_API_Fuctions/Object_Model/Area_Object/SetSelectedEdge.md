# SetSelectedEdge

## Syntax

Sap2000.AreaObj.SetSelectedEdge

## VB6 Procedure

Function SetSelectedEdge(ByVal Name As String, ByVal EdgeNum As Long, ByVal Selected As Boolean) As Long

## Parameters

Name

The name of an existing area object.

EdgeNum

The area object edge that is have its selected status set.

Selected

This item is True if the specified area object edge is selected; otherwise it is False.

## Remarks

This function sets the selected status for area object edges.

The function returns zero if the selected status is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAreaObjectEdgeSelected()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'set area object edge selected

      ret = SapModel.AreaObj.SetSelectedEdge("1", 2, True)



   'update window

      ret = SapModel.View.RefreshWindow



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSelectedEdge](GetSelectedEdge.md)

[GetSelected](GetSelected_%7BArea_Object%7D.md)

[SetSelected](SetSelected_%7BArea_Object%7D.md)
