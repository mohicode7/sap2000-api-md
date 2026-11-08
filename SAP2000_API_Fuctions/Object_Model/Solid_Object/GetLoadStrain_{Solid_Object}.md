# GetLoadStrain

## Syntax

SapObject.SapModel.SolidObj.GetLoadStrain

## VB6 Procedure

Function GetLoadStrain(ByVal Name As String, ByRef NumberItems As Long, ByRef SolidName() As String, ByRef LoadPat() As String, ByRef Component() As Long, ByRef Value() As Double, ByRef PatternName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

NumberItems

The total number of strain loads retrieved for the specified solid objects.

SolidName

This is an array that includes the name of the solid object associated with each strain load.

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

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the solid object specified by the Name item.

If this item is Group, the assignments are retrieved for all solid objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected solid objects, and the Name item is ignored.

## Remarks

This function retrieves the strain load assignments to solid objects.

The function returns zero if the strain load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidObjectStrainLoad()

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



   'get solid object strain load

      ret = SapModel.SolidObj.GetLoadStrain("1", NumberItems, SolidName, LoadPat, Component, Value, PatternName)



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

[SetLoadStrain](SetLoadStrain_%7BSolid_Object%7D.md)

[DeleteLoadStrain](DeleteLoadStrain_%7BSolid_Object%7D.md)
