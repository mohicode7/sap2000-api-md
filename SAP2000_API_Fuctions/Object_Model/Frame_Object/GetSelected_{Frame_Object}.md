# GetSelected

## Syntax

Sap2000.FrameObj.GetSelected

## VB6 Procedure

Function GetSelected(ByVal Name As String, ByRef Selected As Boolean) As Long

## Parameters

Name

The name of an existing frame object.

Selected

This item returns True if the specified frame object is selected, otherwise it returns False.

## Remarks

This function retrieves the selected status for a frame object.

The function returns zero if the selected status is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameObjectSelectedStatus()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set all frames selected

      ret = SapModel.FrameObj.SetSelected("All", True, Group)



   'get frame object selected status

      ret = SapModel.FrameObj.GetSelected("8", Selected)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSelected](SetSelected_%7BFrame_Object%7D.md)
