# GetLoadGravity

## Syntax

SapObject.SapModel.AreaElm.GetLoadGravity

## VB6 Procedure

Function GetLoadGravity(ByVal Name As String, ByRef NumberItems As Long, ByRef AreaName() As String, ByRef LoadPat() As String, ByRef CSys() As String, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing area element or group, depending on the value of the ItemType item.

NumberItems

The total number of gravity loads retrieved for the specified area elements.

AreaName

This is an array that includes the name of the area element associated with each gravity load.

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

If this item is ObjectElm, the load assignments are retrieved for the area elements corresponding to the area object specified by the Name item.

If this item is Element, the load assignments are retrieved for the area element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved for the area elements corresponding to all area objects included in the group specified by the Name item.

If this item is SelectionElm, the load assignments are retrieved for area elements corresponding to all selected area objects, and the Name item is ignored.

## Remarks

This function retrieves the gravity load assignments to area elements.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaElementGravityLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim AreaName() As String

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign auto mesh options

      ret = SapModel.AreaObj.SetAutoMesh("ALL", 1, 3, 3, , , , , , , , , , , , , , , Group)



   'assign area object gravity loads

      ret = SapModel.AreaObj.SetLoadGravity("ALL", "DEAD", 0, 0, -1, , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get area element gravity load

      ret = SapModel.AreaElm.GetLoadGravity("3-1", NumberItems, AreaName, LoadPat, CSys, x, y, z)



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
