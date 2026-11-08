# ChangeCoordinates\_1

## Syntax

SapObject.SapModel.EditPoint.ChangeCoordinates\_1

## VB6 Procedure

Function ChangeCoordinates\_1(ByVal Name As String, ByVal x As Double, ByVal y As Double, ByVal z As Double, Optional ByVal NoRefresh As Boolean = False) As Long

## Parameters

Name

The name of an existing point object.

x, y, z

These are the new x, y and z coordinates, in the present coordinate system, for the specified point object.

NoRefresh

If this item is True, the model display window is not refreshed after the point object is moved.

## Remarks

This function changes the coordinates of a specified point object.

The function returns zero if the coordinate change is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangePointCoordinates_1()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'change point coordinates

      ret = SapModel.EditPoint.ChangeCoordinates_1("1", -288, 0, 36)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.05.

This function supersedes [ChangeCoordinates](../../Obsolete_Functions/ChangeCoordinates_%7BPoint%7D.md).

Modified optional argument NoRefresh to be ByVal in version 12.0.1.

## See Also
