# Connect

## Syntax

SapObject.SapModel.EditPoint.Connect

## VB6 Procedure

Function Connect(ByRef NumberPoints As Long, ByRef PointName() As String) As Long

## Parameters

NumberPoints

The number of the point objects that remain at locations where connections were made.

PointName

This is an array of the name of each point object that remains at locations where connections were made.

## Remarks

This function connects objects that have been disconnected using the Disconnect function. If two or more objects have different end points objects that are at the same location, all of those objects can be connected together by selecting the objects, and selecting their end points, and calling the Connect function. The result will be that all of the objects are connected at a single point.

The function returns zero if the connect is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub ConnectObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim Point1 As String, Point2 As String

      Dim NumberPoints As Long

      Dim PointName() As String



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



   'disconnect point objects

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.PointObj.SetSelected("3", True)

      ret = SapModel.PointObj.SetSelected("9", True)

      ret = SapModel.EditPoint.Disconnect(NumberPoints, PointName)



   'connect objects

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.FrameObj.SetSelected("2", True)

      ret = SapModel.FrameObj.SetSelected("8", True)

      ret = SapModel.FrameObj.GetPoints("2", Point1, Point2)

      ret = SapModel.PointObj.SetSelected(Point2, True)

      ret = SapModel.FrameObj.GetPoints("8", Point1, Point2)

      ret = SapModel.PointObj.SetSelected(Point1, True)

      ret = SapModel.EditPoint.Connect(NumberPoints, PointName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[Disconnect](Disconnect.md)
