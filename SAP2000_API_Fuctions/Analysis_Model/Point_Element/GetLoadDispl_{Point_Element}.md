# GetLoadDispl

## Syntax

SapObject.SapModel.PointElm.GetLoadDispl

## VB6 Procedure

Function GetLoadDispl(ByVal Name As String, ByRef NumberItems As Long, ByRef PointName() As String, ByRef LoadPat() As String, ByRef LCStep() As Long, ByRef CSys() As String, ByRef U1() As Double, ByRef U2() As Double, ByRef U3() As Double, ByRef R1() As Double, ByRef R2() As Double, ByRef R3() As Double, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing point object, point element, or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

This is the total number of joint ground displacement assignments returned.

PointName

This is an array that includes the name of the point element to which the specified ground displacement assignment applies.

LoadPat

This is an array that includes the name of the load pattern for the ground displacement load.

LCStep

This is an array that includes the load pattern step for the ground displacement load. In most cases, this item does not apply and will be returned as 0.

CSys

This is an array that includes the name of the coordinate system for the ground displacement load. This is either Local or the name of a defined coordinate system.

U1

This is an array that includes the assigned translational ground displacement in the local 1-axis or coordinate system X-axis direction, depending on the specified CSys. [L]

U2

This is an array that includes the assigned translational ground displacement in the local 2-axis or coordinate system Y-axis direction, depending on the specified CSys. [L]

U3

This is an array that includes the assigned translational ground displacement in the local 3-axis or coordinate system Z-axis direction, depending on the specified CSys. [L]

R1

This is an array that includes the assigned rotational ground displacement about the local 1-axis or coordinate system X-axis, depending on the specified CSys. [rad]

R2

This is an array that includes the assigned rotational ground displacement about the local 2-axis or coordinate system Y-axis, depending on the specified CSys. [rad]

R3

This is an array that includes the assigned rotational ground displacement about the local 3-axis or coordinate system Z-axis, depending on the specified CSys. [rad]

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the ground displacement assignments are retrieved for the point element corresponding to the point object specified by the Name item.

If this item is Element, the ground displacement assignments are retrieved for the point element specified by the Name item.

If this item is GroupElm, the ground displacement assignments are retrieved for all point elements directly or indirectly specified in the group specified by the Name item.

If this item is SelectionElm, the ground displacement assignments are retrieved for all point elements directly or indirectly selected, and the Name item is ignored.

See [Item Type for Elements](../../Analysis_Results/Results/Item_Type_for_Elements.md) for more information.

## Remarks

This function retrieves the ground displacement load assignments to point elements.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointElmDisplLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim PointName() As String

      Dim LoadPat() As String

      Dim LCStep() As Long

      Dim CSys() As String

      Dim U1() As Double

      Dim U2() As Double

      Dim U3() As Double

      Dim R1() As Double

      Dim R2() As Double

      Dim R3() As Double

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



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get ground displacement load at point elements

      ret = SapModel.PointElm.GetLoadDispl("ALL", NumberItems, PointName, LoadPat, LCStep, CSys, U1, U2, U3, R1, R2, R3, GroupElm)



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
