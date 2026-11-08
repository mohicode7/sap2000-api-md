# GetLoadPoint

## Syntax

SapObject.SapModel.LineElm.GetLoadPoint

## VB6 Procedure

Function GetLoadPoint(ByVal Name As String, ByRef NumberItems As Long, ByRef LineName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef CSys() As String, ByRef Dir() As Long, ByRef RelDist() As Double, ByRef Dist() As Double, ByRef Val() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing line object, line element or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

The total number of point loads retrieved for the specified line elements.

LineName

This is an array that includes the name of the line element associated with each point load.

LoadPat

This is an array that includes the name of the coordinate system in which the point loads are specified.

MyType

This is an array that includes either 1 or 2, indicating the type of point load.

1 = Force

2 = Moment

CSys

This is an array that includes the name of the coordinate system in which each point load is defined. It may be Local or the name of a defined coordinate system.

Dir

This is an array that includes an integer between 1 and 11, indicating the direction of the load.

1 = Local 1 axis (only applies when CSys is Local)

2 = Local 2 axis (only applies when CSys is Local)

3 = Local 3 axis (only applies when CSys is Local)

4 = X direction (does not apply when CSys is Local)

5 = Y direction (does not apply when CSys is Local)

6 = Z direction (does not apply when CSys is Local)

7 = Projected X direction (does not apply when CSys is Local)

8 = Projected Y direction (does not apply when CSys is Local)

9 = Projected Z direction (does not apply when CSys is Local)

10 = Gravity direction (only applies when CSys is Global)

11 = Projected Gravity direction (only applies when CSys is Global)

The positive gravity direction (see Dir = 10 and 11) is in the negative Global Z direction.

RelDist

This is an array that includes the relative distance from the I-End of the line element to the location where the point load is applied.

Dist

This is an array that includes the actual distance from the I-End of the line element to the location where the point load is applied. [L]

Val

This is an array that includes the value of the point load. [F] when MyType is 1 and [FL] when MyType is 2

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the load assignments are retrieved for the line elements corresponding to the line object specified by the Name item.

If this item is Element, the load assignments are retrieved for the line element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved for the line elements corresponding to all line objects included in the group specified by the Name item.

If this item is SelectionElm, the load assignments are retrieved for line elements corresponding to all selected line objects, and the Name item is ignored.

## Remarks

This function retrieves the point load assignments to line elements.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmPointLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim LineName() As String

      Dim LoadPat() As String

      Dim MyType() As Long

      Dim CSys() As String

      Dim Dir() As Long

      Dim RelDist() As Double

      Dim Dist() As Double

      Dim Val() As Double



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



   'assign frame object point loads

      ret = SapModel.FrameObj.SetLoadPoint("14", "DEAD", 1, 10, .5, 20)

      ret = SapModel.FrameObj.SetLoadPoint("15", "DEAD", 1, 10, .5, 20)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element point loads

      ret = SapModel.LineElm.GetLoadPoint("ALL", NumberItems, LineName, LoadPat, MyType, CSys, Dir, RelDist, Dist, Val, GroupElm)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also
