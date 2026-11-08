# Join

## Syntax

SapObject.SapModel.EditFrame.Join

## VB6 Procedure

Function Join(ByVal Name As String, ByVal Item2 As String) As Long

## Parameters

Name

The name of an existing frame object to be joined. The new, joined frame object keeps this name.

Item2

The name of an existing frame object to be joined.

## Remarks

This function joins two straight frame objects that have a common end point and are colinear.

The function returns zero if the frame objects are successfully joined; otherwise it returns a nonzero value.

## VBA Example



```
Sub JoinFrameObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'join frame objects

      ret = SapModel.EditFrame.Join("8", "10")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[DivideAtDistance](DivideAtDistance.md)

[DivideAtIntersections](DivideAtIntersections.md)

[DivideByRatio](DivideByRatio.md)
