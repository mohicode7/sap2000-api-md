# SetLoadDispl

## Syntax

SapObject.SapModel.PointObj.SetLoadDispl

## VB6 Procedure

Function SetLoadDispl(ByVal Name As String, ByVal LoadPat As String, ByRef Value() As Double, Optional ByVal Replace As Boolean = False, Optional ByVal CSys As String = "Global", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

LoadPat

The name of the load pattern for the ground displacement load.

Value

This is an array of six ground displacement load values.

Value(0) = U1 [L]

Value(1) = U2 [L]

Value(2) = U3 [L]

Value(3) = R1 [rad]

Value(4) = R2 [rad]

Value(5) = R3 [rad]

Replace

If this item is True, all previous ground displacement loads, if any, assigned to the specified point object(s) in the specified load pattern are deleted before making the new assignment.

CSys

The name of the coordinate system for the considered ground displacement load. This is Local or the name of a defined coordinate system.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignment is made to the point object specified by the Name item.

If this item is Group, the load assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignment is made to all selected point objects and the Name item is ignored.

## Remarks

This function makes ground displacement load assignments to point objects.

The function returns zero if the load assignments are successfully made, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetPointDisplLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value() As Double



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



   'add ground displacement load

      Redim Value(5)

      Value(0) = 10

      ret = SapModel.PointObj.SetLoadDispl("1", "DEAD", Value)



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

[GetLoadDispl](GetLoadDispl_%7BPoint_Object%7D.md)

[DeleteLoadDispl](DeleteLoadDispl.md)
