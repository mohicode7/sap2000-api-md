# Count

## Syntax

SapObject.SapModel.FrameObj.Count

## VB6 Procedure

Function Count(Optional ByVal MyType As String = "All") As Long

## Parameters

MyType

This is All, Straight, or Curved.

All returns a count of all frame objects in the model, including both straight and curved frame objects. Straight returns a count of all straight frame objects in the model. Curved returns a count of all curved frame objects in the model.

## Remarks

This function returns a count of the frame objects in the model. Depending on the value of the MyType item, the count may be of all frame objects in the model, just the straight frame objects in the model or just the curved frame objects in the model.

## VBA Example



```
Sub CountFrameObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long



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



   'return number of frame objects

      Count = SapModel.FrameObj.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
