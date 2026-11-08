# AssembledJointMass

## Syntax

SapObject.SapModel.Results.AssembledJointMass

## VB6 Procedure

Function AssembledJointMass(ByVal Name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef PointElm() As String, ByRef U1() As Double, ByRef U2() As Double, ByRef U3() As Double, ByRef R1() As Double, ByRef R2() As Double, ByRef R3() As Double) As Long

## Parameters

Name

The name of an existing point element or group of objects, depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the point element corresponding to the point object specified by the Name item.

If this item is Element, the result request is for the point element specified by the Name item.

If this item is GroupElm, the result request is for all point elements directly or indirectly specified in the group specified by the Name item.

If this item is SelectionElm, the result request is for all point elements directly or indirectly selected and the Name item is ignored.

See Item [Type for Elements](../Analysis_Results/Results/Item_Type_for_Elements.md) for more information.

NumberResults

The total number of results returned by the program.

PointElm

This is an array that includes the point element name associated with each result.

U1, U2, U3

These are one dimensional arrays that include the translational mass in the point element local 1, 2 and 3 axes directions, respectively, for each result. [M]

R1, R2, R3

These are one dimensional arrays that include the rotational mass moment of inertia about the point element local 1, 2 and 3 axes, respectively, for each result. [ML2]

## Remarks

This function reports the assembled joint masses for the specified point elements.

The function returns zero if the masses are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](../Analysis_Results/Results/Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetAssembledJointMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberResults As Long

      Dim PointElm() As String

      Dim U1() As Double

      Dim U2() As Double

      Dim U3() As Double

      Dim R1() As Double

      Dim R2() As Double

      Dim R3() As Double



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



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'get assembled joint mass for all point elements

      ret = SapModel.Results.AssembledJointMass("ALL", GroupElm, NumberResults, PointElm, U1, U2, U3, R1, R2, R3)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
