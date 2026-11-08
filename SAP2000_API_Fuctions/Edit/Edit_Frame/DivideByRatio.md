# DivideByRatio

## Syntax

SapObject.SapModel.EditFrame.DivideByRatio

## VB6 Procedure

Function DivideByRatio(ByVal Name As String, ByVal Num As long, ByVal Ratio As Double, ByRef NewName() As String) As Long

## Parameters

Name

The name of an existing straight frame object.

Num

The frame object is divided into this number of new objects.

Ratio

The Last/First length ratio for the new frame objects.

NewName

This is an array that includes the names of the new frame objects.

## Remarks

This function divides straight frame objects based on a specified Last/First length ratio. Curved frame objects are not divided.

The function returns zero if the frame objects are successfully divided; otherwise it returns a nonzero value.

## VBA Example



```
Sub DivideFrameObjectByRatio()

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



   'divide frame object by ratio

      ret = SapModel.EditFrame.DivideByRatio("8", 3, 0.3, NewName)



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
