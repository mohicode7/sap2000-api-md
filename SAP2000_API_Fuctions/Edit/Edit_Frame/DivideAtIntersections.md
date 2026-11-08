# DivideAtIntersections

## Syntax

SapObject.SapModel.EditFrame.DivideAtIntersections

## VB6 Procedure

Function DivideAtIntersections(ByVal Name As String, ByRef Num As Double, ByRef NewName() As String) As Long

## Parameters

Name

The name of an existing straight frame object.

Num

This is the number of frame objects into which the specified frame object is divided.

NewName

This is an array that includes the names of the new frame objects.

## Remarks

This function divides straight frame objects at intersections with selected point objects, line objects, area edges and solid edges. Curved frame objects are not divided.

The function returns zero if the frame objects are successfully divided; otherwise it returns a nonzero value.

## VBA Example



```
Sub DivideFrameObjectAtIntersections()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Num As Long

      Dim NewName() As String

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



   'add point objects to model and select them

      ret = SapModel.PointObj.AddCartesian(-188, 0, 288, Name)

      ret = SapModel.PointObj.SetSelected(Name, True)

      ret = SapModel.PointObj.AddCartesian(-88, 0, 288, Name)

      ret = SapModel.PointObj.SetSelected(Name, True)



   'divide frame object at intersections

      ret = SapModel.EditFrame.DivideAtIntersections("8", Num, NewName)



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

[DivideAtDistance](DivideAtDistance.md)
