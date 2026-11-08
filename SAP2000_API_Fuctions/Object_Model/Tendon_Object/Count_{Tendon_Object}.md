# Count

## Syntax

SapObject.SapModel.TendonObj.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns a count of the tendon objects in the model.

## VBA Example



```
Sub CountTendonObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

      Dim Name As String



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



   'return number of tendon objects

      Count = SapModel.TendonObj.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
