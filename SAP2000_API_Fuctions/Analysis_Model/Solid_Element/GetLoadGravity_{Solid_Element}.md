# GetLoadGravity

## Syntax

SapObject.SapModel.SolidElm.GetLoadGravity

## VB6 Procedure

Function GetLoadGravity(ByVal Name As String, ByRef NumberItems As Long, ByRef SolidName() As String, ByRef LoadPat() As String, ByRef CSys() As String, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing solid element or group, depending on the value of the ItemType item.

NumberItems

The total number of gravity loads retrieved for the specified solid elements.

SolidName

This is an array that includes the name of the solid element associated with each gravity load.

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

If this item is ObjectElm, the load assignments are retrieved for the solid elements corresponding to the solid object specified by the Name item.

If this item is Element, the load assignments are retrieved for the solid element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved for the solid elements corresponding to all solid objects included in the group specified by the Name item.

If this item is SelectionElm,  the load assignments are retrieved for solid elements corresponding to all selected solid objects, and the Name item is ignored.

## Remarks

This function retrieves the gravity load assignments to solid elements.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidElementGravityLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim SolidName() As String

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign auto mesh options

      ret = SapModel.SolidObj.SetAutoMesh("ALL", 1, 2, 2, 2, , , , , , Group)



   'assign solid object gravity loads

      ret = SapModel.SolidObj.SetLoadGravity("ALL", "DEAD", 0, 0, -1, , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get solid element gravity load

      ret = SapModel.SolidElm.GetLoadGravity("3-1", NumberItems, SolidName, LoadPat, CSys, x, y, z)



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
