# SetLoadStrain

## Syntax

SapObject.SapModel.FrameObj.SetLoadStrain

## VB6 Procedure

Function SetLoadStrain(ByVal Name As String, ByVal LoadPat As String, ByVal DOF As Long, ByVal Val As Double, Optional ByVal Replace As Boolean = True, Optional ByVal PatternName As String = "", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

DOF

This is 1, 2, 3, 4, 5 or 6, indicating the degree of freedom to which the strain load is applied.

1 = Strain11

2 = Strain12

3 = Strain13

4 = Curvature1

5 = Curvature2

6 = Curvature3

Val

This is the strain load value. [L/L] for DOF = 1, 2 and 3 and [1/L] for DOF = 4, 5 and 6

Replace

If this item is True, all previous strain loads, if any, assigned to the specified frame object(s), in the specified load pattern, for the specified degree of freedom, are deleted before making the new assignment.

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the strain load for the frame object is uniform along the object at the value specified by Val.

If PatternName is the name of a defined joint pattern, the strain load for the frame object is based on the specified strain value multiplied by the pattern value at the joints at each end of the frame object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns strain loads to frame objects.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameStrainLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'assign frame strain load

      ret = SapModel.FrameObj.SetLoadStrain("1", "DEAD", 1, 0.001)



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

[GetLoadStrain](GetLoadStrain_%7BFrame_Object%7D.md)

[DeleteLoadStrain](DeleteLoadStrain_%7BFrame_Object%7D.md)
