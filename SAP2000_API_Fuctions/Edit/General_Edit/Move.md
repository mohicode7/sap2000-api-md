# Move

## Syntax

SapObject.SapModel.EditGeneral.Move

## VB6 Procedure

Function Move(ByVal dx As Double, ByVal dy As Double, ByVal dz As Double) As Long

## Parameters

dx, dy, dz

These are the x, y and z offsets used, in the present coordinate system, for moving the selected objects.

## Remarks

This function moves selected point, frame, cable, tendon, area, solid and link objects.

The function returns zero if the move is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub MoveObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FrameName() As String



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



   'move objects

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.PointObj.SetSelected("3", True)

      ret = SapModel.PointObj.SetSelected("6", True)

      ret = SapModel.PointObj.SetSelected("9", True)

      ret = SapModel.FrameObj.SetSelected("8", True)

      ret = SapModel.FrameObj.SetSelected("10", True)

      ret = SapModel.EditGeneral.Move(0, 0, 12)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
