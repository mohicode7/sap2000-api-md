# Merge

## Syntax

SapObject.SapModel.EditPoint.Merge

## VB6 Procedure

Function Merge(ByVal MergeTol As Double, ByRef NumberPoints As Long, ByRef PointName() As String) As Long

## Parameters

MergeTol

Point objects within this distance of one another are merged into one point object. [L]

NumberPoints

The number of the selected point objects that still exist after the merge is complete.

PointName

This is an array of the name of each selected point object that still exists after the merge is complete.

## Remarks

This function merges selected point objects that are within a specified distance of one another.

The function returns zero if the merge is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub MergePointObjects()

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



   'add frame object by coordinates

      ret = SapModel.FrameObj.AddByCoord(-400, 0, 288, -289, 0, 288, Name)



   'refresh view

      ret = SapModel.View.RefreshView(0, False)



   'merge point objects

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.PointObj.SetSelected("3", True)

      ret = SapModel.FrameObj.GetPoints(Name, Point1, Point2)

      ret = SapModel.PointObj.SetSelected(Point2, True)

      ret = SapModel.EditPoint.Merge(2, NumberPoints, PointName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
