# GetConstraint

## Syntax

SapObject.SapModel.PointObj.GetConstraint

## VB6 Procedure

Function GetConstraint(ByVal Name As String, ByRef NumberItems As Long, ByRef PointName() As String, ByRef ConstraintName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing point object or group, depending on the value of the ItemType item.

NumberItems

This is the total number of constraint assignments returned.

PointName

This is an array that includes the name of the point object to which the specified constraint assignment applies.

ConstraintName

This is an array that includes the name of the constraint that is assigned to the point object specified by the PointName item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the constraint assignments are retrieved for the point object specified by the Name item.

If this item is Group, the constraint assignments are retrieved for all point objects in the group specified by the Name item.

If this item is SelectedObjects, the constraint assignments are retrieved for all selected point objects and the Name item is ignored.

## Remarks

This function returns a list of constraint assignments made to one or more specified point objects.

The function returns zero if the constraint name list is successfully filled, otherwise it returns nonzero.

The PointName and ConstraintName items are returned in one-dimensional arrays. Each array is created as a dynamic array by the API user. In VBA a dynamic string array is defined by:



```
   Dim PointName() as String
```



The arrays are dimensioned to (NumberItems – 1) inside the Sap2000 program, filled with values, and returned to the API user.

The arrays are zero-based. Thus the first item is at array index 0, and the last item is at array index (NumberItems - 1).

## VBA Example



```
Sub GetConstraintAssignments()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      dim NumberItems as Long

      Dim PointName() As String

      Dim ConstraintName() As String

      Dim i As Long



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



   'define a new constraint

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1")



   'define new constraint assignments

      For i = 4 To 16 Step 4

         ret = SapModel.PointObj.SetConstraint(Format(i), "Diaph1")

      Next i



   'get constraint assignments

      ret = SapModel.PointObj.GetConstraint("ALL", NumberItems, PointName, ConstraintName, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetConstraint](SetConstraint.md)

[DeleteConstraint](DeleteConstraint.md)
