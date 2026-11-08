# PointRemove

## Syntax

SapObject.SapModel.EditArea.PointRemove

## VB6 Procedure

Function PointRemove() As Long

## Parameters

None

## Remarks

This function removes selected point objects from selected area objects. Note that in some cases this command can cause the area object to be deleted.

The function returns zero if it is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub RemovePointFromAreaObject()

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



   'remove point from area object

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.AreaObj.SetSelected("4", True)

      ret = SapModel.PointObj.SetSelected("9", True)

      ret = SapModel.EditArea.PointRemove



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[PointAdd](PointAdd.md)
