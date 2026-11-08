# GetConstraint

## Syntax

SapObject.SapModel.PointElm.GetConstraint

## VB6 Procedure

Function GetConstraint(ByVal Name As String, ByRef NumberItems As Long, ByRef PointName() As String, ByRef ConstraintName() As String, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing point object, point element, or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

This is the total number of constraint assignments returned.

PointName

This is an array that includes the name of the point element to which the specified constraint assignment applies.

ConstraintName

This is an array that includes the name of the constraint that is assigned to the point element specified by the PointName item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm,  the constraint assignments are retrieved for the point element corresponding to the point object specified by the Name item.

If this item is Element, the constraint assignments are retrieved for the point element specified by the Name item.

If this item is GroupElm, the constraint assignments are retrieved for all point elements directly or indirectly specified in the group specified by the Name item.

If this item is SelectionElm, the constraint assignments are retrieved for all point elements directly or indirectly selected, and the Name item is ignored.

See [Item Type for Elements](../../Analysis_Results/Results/Item_Type_for_Elements.md) for more information.

## Remarks

This function returns a list of constraint assignments made to one or more specified point elements.

The function returns zero if the constraint name list is successfully filled, otherwise it returns nonzero.

The PointName and ConstraintName items are returned in one-dimensional arrays. Each array is created as a dynamic array by the API user. In VBA a dynamic string array is defined by:



```
   Dim PointName() as String

 
```



The arrays are dimensioned to (NumberItems – 1) inside the Sap2000 program, filled with values, and returned to the API user.

The arrays are zero-based. Thus the first item is at array index 0, and the last item is at array index (NumberItems - 1).

## VBA Example



```
Sub GetConstraintElmAssignments()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems as Long

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



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get constraint assignments to point elements

      ret = SapModel.PointElm.GetConstraint("ALL", NumberItems, PointName, ConstraintName, GroupElm)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
