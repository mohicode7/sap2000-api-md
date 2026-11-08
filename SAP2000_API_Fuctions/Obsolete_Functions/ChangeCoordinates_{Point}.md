# ChangeCoordinates  (Note:  Newer function available)

## Syntax

SapObject.SapModel.EditPoint.ChangeCoordinates

## VB6 Procedure

Function ChangeCoordinates(ByVal Name As String, ByVal x As Double, ByVal y As Double, ByVal z As Double) As Long

## Parameters

Name

The name of an existing point object.

x, y, z

These are the new x, y and z coordinates, in the present coordinate system, for the specified point object.

## Remarks

This function changes the coordinates of a specified point object.

The function returns zero if the coordinate change is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangePointCoordinates()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'change point coordinates

      ret = SapModel.EditPoint.ChangeCoordinates("1", -288, 0, 36)

      ret = SapModel.View.RefreshWindow



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

The function is obsolete and has been superseded by [ChangeCoordinates\_1](../Edit/Edit_Point/ChangeCoordinates_1_%7BPoint%7D.md) as of version 11.05. This function is maintained for backward compatibility. New function added.

## See Also
