# DivideAtDistance

## Syntax

SapObject.SapModel.EditFrame.DivideAtDistance

## VB6 Procedure

Function DivideAtDistance(ByVal Name As String, ByVal Dist As Double, ByVal IEnd As Boolean, ByRef NewName() As String) As Long

## Parameters

Name

The name of an existing straight frame object.

Dist

The frame object is divided at this distance from the end specified by the IEnd item.[L]

IEnd

If this item is True, the Dist item is measured from the I-end of the frame object. Otherwise it is measured from the J-end of the frame object.

Num

This is the number of frame objects into which the specified frame object is divided.

NewName

This is an array that includes the names of the two new frame objects.

## Remarks

This function divides straight frame objects into two objects at a location defined by the Dist and IEnd items. Curved frame objects are not divided.

The function returns zero if the frame objects are successfully divided; otherwise it returns a nonzero value.

## VBA Example



```
Sub DivideFrameObjectAtDistance()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NewName() As String



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



   'divide frame object at distance

      ret = SapModel.EditFrame.DivideAtDistance("8", 100, True, NewName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[DivideByRatio](DivideByRatio.md)

[DivideAtIntersections](DivideAtIntersections.md)
