# PreviousSelection

## Syntax

SapObject.SapModel.SelectObj.PreviousSelection

## VB6 Procedure

Function PreviousSelection() As Long

## Parameters

None

## Remarks

This function restores the previous selection.

The function returns zero if the selection is successfully restored, otherwise it returns nonzero.

## VBA Example



```
Sub RestoreSelection()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'select objects

      ret = SapModel.SelectObj.PlaneXY("2")



   'clear selection

      ret = SapModel.SelectObj.ClearSelection



   'get previous selection

      ret = SapModel.SelectObj.PreviousSelection



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
