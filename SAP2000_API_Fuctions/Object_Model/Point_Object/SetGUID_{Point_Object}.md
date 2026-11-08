# SetGUID

## Syntax

SapObject.SapModel.PointObj.SetGUID

## VB6 Procedure

Function SetGUID(ByVal Name As String, Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing point object.

GUID

The GUID (Global Unique ID) for the specified point object.

## Remarks

This function sets the GUID for specified point object. If the GUID is passed in as a blank string, the program automatically creates a GUID for the object.

This function returns zero if the point object GUID is successfully set; otherwise it returns nonzero.

## VBA Example



```
Sub SetPointGUID()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'set program created GUID

      ret = SapModel.PointObj.SetGUID("1")



   'get GUID

      ret = SapModel.PointObj.GetGUID("1", GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetGUID](GetGUID_%7BPoint_Object%7D.md)
