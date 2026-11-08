# GetLoadDistributed

## Syntax

SapObject.SapModel.LineElm.GetLoadDistributed

## VB6 Procedure

Function GetLoadDistributed(ByVal Name As String, ByRef NumberItems As Long, ByRef LineName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef CSys() As String, ByRef Dir() As Long, ByRef RD1() As Double, ByRef RD2() As Double, ByRef Dist1() As Double, ByRef Dist2() As Double, ByRef Val1() As Double, ByRef Val2() As Double, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing line object, line element or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

The total number of distributed loads retrieved for the specified line elements.

LineName

This is an array that includes the name of the line element associated with each distributed load.

LoadPat

This is an array that includes the name of the coordinate system in which the distributed loads are specified.

MyType

This is an array that includes either 1 or 2, indicating the type of distributed load.

1 = Force

2 = Moment

CSys

This is an array that includes the name of the coordinate system in which each distributed load is defined. It may be Local or the name of a defined coordinate system.

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

RD1

This is an array that includes the relative distance from the I-End of the line element to the start of the distributed load.

RD2

This is an array that includes the relative distance from the I-End of the line element to the end of the distributed load.

Dist1

This is an array that includes the actual distance from the I-End of the line element to the start of the distributed load. [L]

Dist2

This is an array that includes the actual distance from the I-End of the line element to the end of the distributed load. [L]

Val1

This is an array that includes the load value at the start of the distributed load. [F/L] when MyType is 1 and [FL/L] when MyType is 2

Val2

This is an array that includes the load value at the end of the distributed load. [F/L] when MyType is 1 and [FL/L] when MyType is 2

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

This function retrieves the distributed load assignments to line elements.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmDistributedLoad()

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

      Dim RD1() As Double

      Dim RD2() As Double

      Dim Dist1() As Double

      Dim Dist2() As Double

      Dim Val1() As Double

      Dim Val2() As Double



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



   'assign frame object distributed loads

      ret = SapModel.FrameObj.SetLoadDistributed("14", "DEAD", 1, 10, 0, 1, 0.08, 0.04)



   'assign frame object auto mesh options

      ret = SapModel.FrameObj.SetAutoMesh("ALL", True, False, False, 2, 0, Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element distributed loads

      ret = SapModel.LineElm.GetLoadDistributed("14-1", NumberItems, LineName, LoadPat, MyType, CSys, Dir, RD1, RD2, Dist1, Dist2, Val1, Val2)



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
