# GetLoadTemperature

## Syntax

SapObject.SapModel.LineElm.GetLoadTemperature

## VB6 Procedure

Function GetLoadTemperature(ByVal Name As String, ByRef NumberItems As Long, ByRef LineName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef Val() As Double, ByRef PatternName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing line object, line element or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

The total number of temperature loads retrieved for the specified line elements.

LineName

This is an array that includes the name of the line element associated with each temperature load.

LoadPat

This is an array that includes the name of the load pattern associated with each temperature load.

MyType

This is an array that includes 1, 2 or 3, indicating the type of temperature load.

1 = Temperature

2 = Temperature gradient along local 2 axis

3 = Temperature gradient along local 3 axis

Val

This is an array that includes the temperature load value. [T] for MyType= 1 and [T/L] for MyType= 2 and 3

PatternName

This is an array that includes the joint pattern name, if any, used to specify the temperature load.

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

This function retrieves the temperature load assignments to line elements.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmTemperatureLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim LineName() As String

      Dim LoadPat() As String

      Dim MyType() As Long

      Dim Val() As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign frame object temperature load

      ret = SapModel.FrameObj.SetLoadTemperature("All", "DEAD", 1, 50, , , Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element temperature load

      ret = SapModel.LineElm.GetLoadTemperature("ALL", NumberItems, LineName, LoadPat, MyType, Val, PatternName, GroupElm)



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
