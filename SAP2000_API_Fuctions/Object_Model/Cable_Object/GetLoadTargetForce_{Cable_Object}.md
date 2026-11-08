# GetLoadTargetForce

## Syntax

SapObject.SapModel.CableObj.GetLoadTargetForce

## VB6 Procedure

Function GetLoadTargetForce(ByVal Name As String, ByRef NumberItems As Long, ByRef CableName() As String, ByRef LoadPat() As String, ByRef P() As Double, ByRef RD() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

NumberItems

The total number of deformation loads retrieved for the specified cable objects.

CableName

This is an array that includes the name of the cable object associated with each target force.

LoadPat

This is an array that includes the name of the load pattern associated with each target force.

P

This is an array of axial target force values. [F]

RD

This is an array of the relative distances along the cable objects where the axial target force values apply.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the cable object specified by the Name item.

If this item is Group, the assignments are retrieved for all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected cable objects, and the Name item is ignored.

## Remarks

This function retrieves the target force assignments to cable objects.

The function returns zero if the target force assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableTargetForce()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim CableName() As String

      Dim LoadPat() As String

      Dim P() As double

      Dim RD() As double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'assign cable target force

      ret = SapModel.CableObj.SetLoadTargetForce(Name, "DEAD", 50, 0.5)



   'get cable target force

      ret = SapModel.CableObj.GetLoadTargetForce(Name, NumberItems, CableName, LoadPat, P, RD)



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

[SetLoadTargetForce](SetLoadTargetForce_%7BCable_Object%7D.md)

[DeleteLoadTargetForce](DeleteLoadTargetForce_%7BCable_Object%7D.md)
