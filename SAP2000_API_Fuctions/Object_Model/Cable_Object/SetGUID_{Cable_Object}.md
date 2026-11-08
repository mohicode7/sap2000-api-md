# SetGUID

## Syntax

SapObject.SapModel.CableObj.SetGUID

## VB6 Procedure

Function SetGUID(ByVal Name As String, Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing cable object.

GUID

The GUID (Global Unique ID) for the specified cable object.

## Remarks

This function sets the GUID for the specified cable object. If the GUID is passed in as a blank string, the program automatically creates a GUID for the object.

This function returns zero if the cable object GUID is successfully set; otherwise it returns nonzero.

## VBA Example



```
Sub SetCableObjGUID()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'set program created GUID

      ret = SapObject.SapModel.CableObj.SetGUID(Name)



   'get GUID

      ret = SapObject.SapModel.CableObj.GetGUID(Name, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetGUID](GetGUID_%7BCable_Object%7D.md)
