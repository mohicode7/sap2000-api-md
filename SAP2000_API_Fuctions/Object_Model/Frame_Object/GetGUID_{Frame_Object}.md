# GetGUID

## Syntax

SapObject.SapModel.FrameObj.GetGUID

## VB6 Procedure

Function GetGUID(ByVal name As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing frame object.

GUID

The GUID (Global Unique ID) for the specified frame object.

## Remarks

This function retrieves the GUID for the specified frame object.

This function returns zero if the frame object GUID is successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetFrameObjGUID()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim GUID As String



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



   'set program created GUID

      ret = SapObject.SapModel.FrameObj.SetGUID("1")



   'get GUID

      ret = SapObject.SapModel.FrameObj.GetGUID("1", GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetGUID](SetGUID_%7BFrame_Object%7D.md)
