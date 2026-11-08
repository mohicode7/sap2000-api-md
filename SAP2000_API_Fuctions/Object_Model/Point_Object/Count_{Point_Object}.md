# Count

## Syntax

SapObject.SapModel.PointObj.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the total number of point objects in the model.

## VBA Example



```
Sub CountPointObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim Count as Long



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



   'return number of point objects

      Count = SapModel.PointObj.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
