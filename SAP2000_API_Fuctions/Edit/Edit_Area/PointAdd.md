# PointAdd

## Syntax

SapObject.SapModel.EditArea.PointAdd

## VB6 Procedure

Function PointAdd() As Long

## Parameters

None

## Remarks

This function adds a point object at the midpoint of selected area object edges.

The function returns zero if it is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddPointToAreaObject()

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



   'add point to area object

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.AreaObj.SetSelectedEdge("4", 2, True)

      ret = SapModel.EditArea.PointAdd



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[PointRemove](PointRemove.md)
