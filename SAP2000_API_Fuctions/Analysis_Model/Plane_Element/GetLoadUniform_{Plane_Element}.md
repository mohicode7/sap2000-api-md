# GetLoadUniform

## Syntax

SapObject.SapModel.PlaneElm.GetLoadUniform

## VB6 Procedure

Function GetLoadUniform(ByVal Name As String, ByRef NumberItems As Long, ByRef PlaneName() As String, ByRef LoadPat() As String, ByRef CSys() As String, ByRef Dir() As Long, ByRef Value() As Double, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing plane element or group, depending on the value of the ItemType item.

NumberItems

The total number of uniform loads retrieved for the specified plane elements.

PlaneName

This is an array that includes the name of the plane element associated with each uniform load.

LoadPat

This is an array that includes the name of the coordinate system in which the uniform load is specified.

CSys

This is an array that includes the name of the coordinate system associated with each uniform load.

Dir

This is an integer between 1 and 11, indicating the direction of the load.

1 = Local 1 axis (applies only when CSys is Local)

2 = Local 2 axis (applies only when CSys is Local)

3 = Local 3 axis (applies only when CSys is Local)

4 = X direction (does not apply when CSys is Local)

5 = Y direction (does not apply when CSys is Local)

6 = Z direction (does not apply when CSys is Local)

7 = Projected X direction (does not apply when CSys is Local)

8 = Projected Y direction (does not apply when CSys is Local)

9 = Projected Z direction (does not apply when CSys is Local)

10 = Gravity direction (applies only when CSys is Global)

11 = Projected Gravity direction (applies only when CSys is Global)

The positive gravity direction (see Dir = 10 and 11) is in the negative Global Z direction.

Value

The uniform load value. [F/L2]

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the load assignments are retrieved for the plane elements corresponding to the area object specified by the Name item.

If this item is Element, the load assignments are retrieved for the plane element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved for the plane elements corresponding to all area objects included in the group specified by the Name item.

If this item is SelectionElm, the load assignments are retrieved for plane elements corresponding to all selected area objects, and the Name item is ignored.

## Remarks

This function retrieves the uniform load assignments to plane elements.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPlaneElementUniformLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim PlaneName() As String

      Dim LoadPat() As String

      Dim CSys() As String

      Dim Dir() As Long

      Dim Value() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 3-001-incomp.sdb")



   'assign area object uniform loads

      ret = SapModel.AreaObj.SetLoadUniform("ALL", "Membrane", -0.01, 2, False, "Local", Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get plane element uniform load

      ret = SapModel.PlaneElm.GetLoadUniform("3", NumberItems, PlaneName, LoadPat, CSys, Dir, Value)



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
