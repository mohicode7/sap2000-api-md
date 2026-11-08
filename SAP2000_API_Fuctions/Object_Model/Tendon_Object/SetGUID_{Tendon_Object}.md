# SetGUID

## Syntax

SapObject.SapModel.TendonObj.SetGUID

## VB6 Procedure

Function SetGUID(ByVal Name As String, Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing tendon object.

GUID

The GUID (Global Unique ID) for the specified tendon object.

## Remarks

This function sets the GUID for the specified tendon object. If the GUID is passed in as a blank string, the program automatically creates a GUID for the object.

This function returns zero if the tendon object GUID is successfully set; otherwise it returns nonzero.

## VBA Example



```
Sub SetTendonObjGUID()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

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



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'set program created GUID

      ret = SapObject.SapModel.TendonObj.SetGUID(Name)



   'get GUID

      ret = SapObject.SapModel.TendonObj.GetGUID(Name, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetGUID](GetGUID_%7BTendon_Object%7D.md)
