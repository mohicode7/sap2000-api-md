# GetLoadStrain

## Syntax

SapObject.SapModel.PlaneElm.GetLoadStrain

## VB6 Procedure

Function GetLoadStrain(ByVal Name As String, ByRef NumberItems
As Long, ByRef PlaneName() As String, ByRef LoadPat() As String, ByRef
Component() As Long, ByRef Value() As Double, ByRef PatternName() As String,
Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing plane element or group, depending
on the value of the ItemType item.

NumberItems

The total number of strain loads retrieved for the specified
plane elements.

PlaneName

This is an array that includes the name of the plane
element associated with each strain load.

LoadPat

This is an array that includes the name of the load
pattern associated with each strain load.

Component

This is an array that includes 1, 2, 3, 4, 5 or 6, indicating
the component associated with each strain load.

1 = Strain11

2 = Strain22

3 = Strain12

4 = Strain13

5 = Strain23

6 = Strain33

Value

This is an array that includes the strain value. [L/L]

PatternName

This is an array that includes the joint pattern name,
if any, used to specify the strain load.

ItemTypeElm

This is one of the following items in the eItemTypeElm
enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the load assignments are
retrieved for the plane elements corresponding to the area object specified
by the Name item.

If this item is Element, the load assignments are retrieved
for the plane element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved
for the plane elements corresponding to all area objects included in the
group specified by the Name item.

If this item is SelectionElm, the load assignments are
retrieved for plane elements corresponding to all selected area objects,
and the Name item is ignored.

## Remarks

This function retrieves the strain load assignments
to plane elements.

The function returns zero if the strain load assignments
are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPlaneElementStrainLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim PlaneName() As String

      Dim LoadPat() As String

      Dim Component() As Long

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



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 
 3-001-incomp.sdb")



   'assign area object strain load

      ret = SapModel.AreaObj.SetLoadStrain("ALL", 
 "Membrane", 1, 0.001, , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get plane element strain load

      ret = SapModel.PlaneElm.GetLoadStrain("3", 
 NumberItems, PlaneName, LoadPat, Component, Value, PatternName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases
and Response Combinations to Load Patterns, Load Cases and Load Combinations,
respectively, in version 12.00.

Added strain in the 33 direction for version 24.1.0.

## See Also
