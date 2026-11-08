# GetLoadGravity

## Syntax

SapObject.SapModel.LineElm.GetLoadGravity

## VB6 Procedure

Function GetLoadGravity(ByVal Name As String, ByRef NumberItems As Long, ByRef LineName() As String, ByRef LoadPat() As String, ByRef CSys() As String, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing line object, line element or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

The total number of gravity loads retrieved for the specified line elements.

LineName

This is an array that includes the name of the line element associated with each gravity load.

LoadPat

This is an array that includes the name of the coordinate system in which the gravity load multipliers are specified.

CSys

This is an array that includes the name of the coordinate system associated with each gravity load.

x, y, z

These are arrays of gravity load multipliers in the x, y and z directions of the specified coordinate system.

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

This function retrieves the gravity load assignments to line elements.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmGravityLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim LineName() As String

      Dim LoadPat() As String

      Dim CSys() As String

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double



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



   'assign frame object gravity loads

      ret = SapModel.FrameObj.SetLoadGravity("ALL", "DEAD", 0, 0, -1, , , Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element gravity load

      ret = SapModel.LineElm.GetLoadGravity("3-1", NumberItems, LineName, LoadPat, CSys, x, y, z)



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
