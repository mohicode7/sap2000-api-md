# GetTypeOAPI

## Syntax

SapObject.SapModel.FrameObj.GetTypeOAPI

## VB6 Procedure

Function GetTypeOAPI(ByVal Name As String, ByRef MyType As String) As Long

## Parameters

Name

The name of a defined frame object.

MyType

This is Straight or Curved, indicating the type of frame object.

## Remarks

This function retrieves the type of frame object (straight or curved).

The function returns zero if the frame object type is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As String



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



   'get frame type

      ret = SapModel.FrameObj.GetTypeOAPI("3", MyType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed function name to GetTypeOAPI in v17.0.0.

## See Also

[SetStraight](SetStraight.md)

[SetCurved](SetCurved.md)

[GetCurved](../../Obsolete_Functions/GetCurved.md)
