# GetSelectedEdge

## Syntax

Sap2000.AreaObj.GetSelectedEdge

## VB6 Procedure

Function GetSelectedEdge(ByVal Name As String, ByRef NumberEdges As Long, ByRef Selected() As Boolean) As Long

## Parameters

Name

The name of an existing area object.

NumberEdges

The number of edges in the specified area object.

Selected

This is an array of items that is True if the specified area object edge is selected; otherwise it is False.

Selected(0) = Selected status for edge 1

Selected(1) = Selected status for edge 2

Selected(n) = Selected status for edge (n + 1)

This array is internally dimensioned by Sap2000 to (NumberEdges – 1).

## Remarks

This function retrieves the selected status for area object edges.

The function returns zero if the selected status is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectEdgesSelectedStatus()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberEdges As Long

      Dim Selected() As Boolean



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

      ret = SapModel.AreaObj.SetSelectedEdge("1", 3, True)



   'get area object edge selected status

      ret = SapModel.AreaObj.GetSelectedEdge("1", NumberEdges, Selected)



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

[GetSelected](GetSelected_%7BArea_Object%7D.md)

[SetSelectedEdge](SetSelectedEdge.md)
