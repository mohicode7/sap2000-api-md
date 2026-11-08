# GetSelected

## Syntax

SapObject.SapModel.PointObj.GetSelected

## VB6 Procedure

Function GetSelected(ByVal Name As String, ByRef Selected As Boolean) As Long

## Parameters

Name

The name of an existing point object.

Selected

This item returns True if the specified point object is selected, otherwise it returns False.

## Remarks

This function retrieves the selected status for a point object.

The function returns zero if the selected status is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointSelectedStatus()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'set point selected

      ret = SapModel.PointObj.SetSelected("3", True)



   'get point selected status

      ret = SapModel.PointObj.GetSelected("3", Selected)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSelected](SetSelected_%7BPoint_Object%7D.md)
