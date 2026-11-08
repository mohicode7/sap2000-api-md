# GetLoadDeformation

## Syntax

SapObject.SapModel.LinkElm.GetLoadDeformation

## VB6 Procedure

Function GetLoadDeformation(ByVal Name As String, ByRef NumberItems As Long, ByRef LinkName() As String, ByRef LoadPat() As String, ByRef dof1() As Boolean, ByRef dof2() As Boolean, ByRef dof3() As Boolean, ByRef dof4() As Boolean, ByRef dof5() As Boolean, ByRef dof6() As Boolean, ByRef U1() As Double, ByRef U2() As Double, ByRef U3() As Double, ByRef R1() As Double, ByRef R2() As Double, ByRef R3() As Double, Optional ByVal ItemTypeElm As eItemTypeElm = Element) As Long

## Parameters

Name

The name of an existing link object, link element or group of objects, depending on the value of the ItemTypeElm item.

NumberItems

The total number of deformation loads retrieved for the specified link elements.

LinkName

This is an array that includes the name of the link element associated with each deformation load.

LoadPat

This is an array that includes the name of the load pattern associated with each deformation load.

dof1, dof2, dof3, dof4, dof5, dof6

These are arrays of boolean values, indicating if the considered degree of freedom has a deformation load.

dof1 = U1

dof2 = U2

dof3 = U3

dof4 = R1

dof5 = R2

dof6 = R3

U1, U2, U3, R1, R2, R3

These are arrays of deformation load values. The deformations specified for a given degree of freedom are applicable only if the corresponding DOF item for that degree of freedom is True.

U1 = U1 deformation [L]

U2 = U2 deformation [L]

U3 = U3 deformation [L]

R1 = R1 deformation [rad]

R2 = R2 deformation [rad]

R3 = R3 deformation [rad]

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the load assignments are retrieved for the link elements corresponding to the link object specified by the Name item.

If this item is Element, the load assignments are retrieved for the link element specified by the Name item.

If this item is GroupElm, the load assignments are retrieved for the link elements corresponding to all link objects included in the group specified by the Name item.

If this item is SelectionElm, the load assignments are retrieved for link elements corresponding to all selected link objects, and the Name item is ignored.

## Remarks

This function retrieves the deformation load assignments to link elements.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkElmDeformationLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim d() As double

      Dim NumberItems As Long

      Dim LinkName() As String

      Dim LoadPat() As String

      Dim dof1() As Boolean

      Dim dof2() As Boolean

      Dim dof3() As Boolean

      Dim dof4() As Boolean

      Dim dof5() As Boolean

      Dim dof6() As Boolean

      Dim U1() As Double

      Dim U2() As Double

      Dim U3() As Double

      Dim R1() As Double

      Dim R2() As Double

      Dim R3() As Double

      Dim Name1 As String

      Dim Name2 As String



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



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("3", "", Name1, True)

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name2)



   'refresh view

      ret = SapModel.View.RefreshView



   'assign link object deformation loads

      ReDim DOF(5)

      ReDim d(5)

      DOF(0) = True

      D(0) = 2

      ret = SapModel.LinkObj.SetLoadDeformation("ALL", "DEAD", DOF, d, Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get link element deformation loads

      ret = SapModel.LinkElm.GetLoadDeformation("ALL", NumberItems, LinkName, LoadPat, dof1, dof2, dof3, dof4, dof5, dof6, U1, U2, U3, R1, R2, R3, GroupElm)



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
