# SetGUID

## Syntax

SapObject.SapModel.AreaObj.SetGUID

## VB6 Procedure

Function SetGUID(ByVal Name As String, Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing area object.

GUID

The GUID (Global Unique ID) for the specified area object.

## Remarks

This function sets the GUID for the specified area object. If the GUID is passed in as a blank string, the program automatically creates a GUID for the object.

This function returns zero if the area object GUID is successfully set; otherwise, it returns nonzero.

## VBA Example



```
Sub SetAreaObjGUID()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'set program created GUID

      ret = SapObject.SapModel.AreaObj.SetGUID("1")



   'get GUID

      ret = SapObject.SapModel.AreaObj.GetGUID("1", GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetGUID](GetGUID_%7BArea_Object%7D.md)
