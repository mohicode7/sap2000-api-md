# GetLoadStrain

## Syntax

SapObject.SapModel.TendonObj.GetLoadStrain

## VB6 Procedure

Function GetLoadStrain(ByVal Name As String, ByRef NumberItems As Long, ByRef TendonName() As String, ByRef LoadPat() As String, ByRef Strain() As Double, ByRef PatternName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing tendon object or group, depending on the value of the ItemType item.

NumberItems

The total number of strain loads retrieved for the specified tendon objects.

TendonName

This is an array that includes the name of the tendon object associated with each strain load.

LoadPat

This is an array that includes the name of the load pattern associated with each strain load.

Strain

This is an array that includes the axial strain value. [L/L]

PatternName

This is an array that includes the joint pattern name, if any, used to specify the strain load.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the tendon object specified by the Name item.

If this item is Group, the assignments are retrieved for all tendon objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected tendon objects, and the Name item is ignored.

## Remarks

This function retrieves the strain load assignments to tendon objects.

The function returns zero if the strain load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonStrainLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim TendonName() As String

      Dim LoadPat() As String

      Dim Strain() As Double

      Dim PatternName() As String

      Dim Name As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'assign tendon strain load

      ret = SapModel.TendonObj.SetLoadStrain(Name, "DEAD", 0.001)



   'get tendon strain load

      ret = SapModel.TendonObj.GetLoadStrain(Name, NumberItems, TendonName, LoadPat, Strain, PatternName)



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

[SetLoadStrain](SetLoadStrain_%7BTendon_Object%7D.md)

[DeleteLoadStrain](DeleteLoadStrain_%7BTendon_Object%7D.md)
