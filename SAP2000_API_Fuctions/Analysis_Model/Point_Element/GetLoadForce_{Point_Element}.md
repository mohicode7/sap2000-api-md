# GetLoadForce

## Syntax

SapObject.SapModel.PointElm.GetLoadForce

## VB6 Procedure

Function GetLoadForce(ByVal Name As String, ByRef NumberItems As Long, ByRef PointName() As String, ByRef LoadPat() As String, ByRef LCStep() As Long, ByRef CSys() As String, ByRef F1() As Double, ByRef F2() As Double, ByRef F3() As Double, ByRef M1() As Double, ByRef M2() As Double, ByRef M3() As Double, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing point object, point element, or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

This is the total number of joint force load assignments returned.

PointName

This is an array that includes the name of the point element to which the specified load assignment applies.

LoadPat

This is an array that includes the name of the load pattern for the load.

LCStep

This is an array that includes the load pattern step for the load. In most cases this item does not apply and will be returned as 0.

CSys

This is an array that includes the name of the coordinate system for the load. This is either Local or the name of a defined coordinate system.

F1

This is an array that includes the assigned translational force in the local 1-axis or coordinate system X-axis direction, depending on the specified CSys. [F]

F2

This is an array that includes the assigned translational force in the local 2-axis or coordinate system Y-axis direction, depending on the specified CSys. [F]

F3

This is an array that includes the assigned translational force in the local 3-axis or coordinate system Z-axis direction, depending on the specified CSys. [F]

M1

This is an array that includes the assigned moment about the local 1-axis or coordinate system X-axis, depending on the specified CSys. [FL]

M2

This is an array that includes the assigned moment about the local 2-axis or coordinate system Y-axis, depending on the specified CSys. [FL]

M3

This is an array that includes the assigned moment about the local 3-axis or coordinate system Z-axis, depending on the specified CSys. [FL]

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the load assignments are retrieved for the point element corresponding to the point object specified by the Name item.

If this item is Element, the load assignments are retrieved for the point element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved for all point elements directly or indirectly specified in the group specified by the Name item.

If this item is SelectionElm, the load assignments are retrieved for all point elements directly or indirectly selected and the Name item is ignored.

See [Item Type for Elements](../../Analysis_Results/Results/Item_Type_for_Elements.md) for more information.

## Remarks

This function retrieves the joint force load assignments to point elements.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointElmForceLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim PointName() As String

      Dim LoadPat() As String

      Dim LCStep() As Long

      Dim CSys() As String

      Dim F1() As Double

      Dim F2() As Double

      Dim F3() As Double

      Dim M1() As Double

      Dim M2() As Double

      Dim M3() As Double

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



   'add joint force load

      Redim Value(5)

      Value(0) = 10

      ret = SapModel.PointObj.SetLoadForce("1", "DEAD", Value)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get point element force load

      ret = SapModel.PointElm.GetLoadForce("ALL", NumberItems, PointName, LoadPat, LCStep, CSys, F1, F2, F3, M1, M2, M3, GroupElm)



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
