# Disconnect

## Syntax

SapObject.SapModel.EditPoint.Disconnect

## VB6 Procedure

Function Disconnect(ByRef NumberPoints As Long, ByRef PointName() As String) As Long

## Parameters

NumberPoints

The number of the point objects (including the original selected point objects) that are created by the disconnect action.

PointName

This is an array of the name of each point object (including the original selected point objects) that is created by the disconnect action.

## Remarks

This function disconnects selected point objects. Disconnect creates a separate point for each object that frames into the selected point object.

The function returns zero if the disconnect is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub DisconnectPointObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberPoints As Long

      Dim PointName() As String



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



   'disconnect point objects

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.PointObj.SetSelected("3", True)

      ret = SapModel.EditPoint.Disconnect(NumberPoints, PointName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[Connect](Connect.md)
