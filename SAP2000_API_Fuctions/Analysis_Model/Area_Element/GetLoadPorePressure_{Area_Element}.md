# GetLoadPorePressure

## Syntax

SapObject.SapModel.AreaElm.GetLoadPorePressure

## VB6 Procedure

Function GetLoadPorePressure(ByVal Name As String, ByRef NumberItems As Long, ByRef AreaName() As String, ByRef LoadPat() As String, ByRef Value() As Double, ByRef PatternName() As String, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing area element or group depending on the value of the ItemType item.

NumberItems

The total number of pore pressure loads retrieved for the specified area elements.

AreaName

This is an array that includes the name of the area element associated with each pore pressure load.

LoadPat

This is an array that includes the name of the load pattern associated with each pore pressure load.

Value

This is an array that includes the pore pressure load value. [F/L2]

PatternName

This is an array that includes the joint pattern name, if any, used to specify the pore pressure load.

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

This function retrieves the pore pressure load assignments to area elements.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaElementPorePressureLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim AreaName() As String

      Dim LoadPat() As String

      Dim Value() As Double

      Dim PatternName() As String



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



   'assign area object pore pressure load

      ret = SapModel.AreaObj.SetLoadPorePressure("ALL", "DEAD", .1, , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get area element pore pressure load

      ret = SapModel.AreaElm.GetLoadPorePressure("ALL", NumberItems, AreaName, LoadPat, Value, PatternName, GroupElm)



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
