# GetLoadStrain

## Syntax

SapObject.SapModel.SolidElm.GetLoadStrain

## VB6 Procedure

Function GetLoadStrain(ByVal Name As String, ByRef NumberItems As Long, ByRef SolidName() As String, ByRef LoadPat() As String, ByRef Component() As Long, ByRef Value() As Double, ByRef PatternName() As String, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing solid element or group, depending on the value of the ItemType item.

NumberItems

The total number of strain loads retrieved for the specified solid elements.

SolidName

This is an array that includes the name of the solid element associated with each strain load.

LoadPat

This is an array that includes the name of the load pattern associated with each strain load.

Component

This is 1, 2, 3, 4, 5 or 6, indicating the component to which the strain load is applied.

1 = Strain11

2 = Strain22

3 = Strain33

4 = Strain12

5 = Strain13

6 = Strain23

Value

This is an array that includes the strain value. [L/L]

PatternName

This is an array that includes the joint pattern name, if any, used to specify the strain load.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the load assignments are retrieved for the solid elements corresponding to the solid object specified by the Name item.

If this item is Element, the load assignments are retrieved for the solid element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved for the solid elements corresponding to all solid objects included in the group specified by the Name item.

If this item is SelectionElm, the load assignments are retrieved for solid elements corresponding to all selected solid objects, and the Name item is ignored.

## Remarks

This function retrieves the strain load assignments to solid elements.

The function returns zero if the strain load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidElementStrainLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim SolidName() As String

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



   'create model from template

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign solid object strain load

      ret = SapModel.SolidObj.SetLoadStrain("ALL", "DEAD", 1, 0.001, , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get solid element strain load

      ret = SapModel.SolidElm.GetLoadStrain("1", NumberItems, SolidName, LoadPat, Component, Value, PatternName)



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
