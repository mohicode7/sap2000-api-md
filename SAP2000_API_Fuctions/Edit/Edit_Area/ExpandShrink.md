# ExpandShrink

## Syntax

SapObject.SapModel.EditArea.ExpandShrink

## VB6 Procedure

Function Divide(ByVal OffsetType As Long, ByVal Offset As Double) As Long

## Parameters

OffsetType

This item is 0, 1 or 2, indicating the offset type for the selected area objects.

0 = Offset all area edges

1 = Offset selected area edges only

2 = Offset selected points of selected areas only

Offset

The area edge offset distance. Positive distances expand the object and negative distances shrink the object.[L]

## Remarks

This function expands or shrinks selected area objects.

The function returns zero if it is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub ExpandAreaObject()

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



   'expand area object

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.AreaObj.SetSelectedEdge("4", 2, True)

      ret = SapModel.EditArea.ExpandShrink(1, 48)



   'refresh view, updating zoom

      ret = SapModel.View.RefreshView(0, False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
