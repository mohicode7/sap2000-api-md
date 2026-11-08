# GetGUID

## Syntax

SapObject.SapModel.SolidObj.GetGUID

## VB6 Procedure

Function GetGUID(ByVal name As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing solid object.

GUID

The GUID (Global Unique ID) for the specified solid object.

## Remarks

This function retrieves the GUID for the specified solid object.

This function returns zero if the solid object GUID is successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetSolidObjGUID()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'set program created GUID

      ret = SapObject.SapModel.SolidObj.SetGUID("1")



   'get GUID

      ret = SapObject.SapModel.SolidObj.GetGUID("1", GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetGUID](SetGUID_%7BSolid_Object%7D.md)
