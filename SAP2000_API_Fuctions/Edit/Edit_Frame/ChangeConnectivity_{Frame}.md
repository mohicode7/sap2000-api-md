# ChangeConnectivity

## Syntax

SapObject.SapModel.EditFrame.ChangeConnectivity

## VB6 Procedure

Function ChangeConnectivity(ByVal Name As String, ByVal Point1 As String, ByVal Point2 As String) As Long

## Parameters

Name

The name of an existing frame object.

Point1

The name of the point object at the I-End of the frame object.

Point2

The name of the point object at the J-End of the frame object.

## Remarks

This function modifies the connectivity of a frame object.

The function returns zero if the frame object connectivity is successfully modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub ModifyFrameObjConnectivity()

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



   'modify connectivity

      ret = SapModel.EditFrame.ChangeConnectivity("8", "3", "5")

      ret = SapModel.View.RefreshWindow



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also
