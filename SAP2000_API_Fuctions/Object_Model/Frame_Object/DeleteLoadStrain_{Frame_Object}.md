# DeleteLoadStrain

## Syntax

SapObject.SapModel.FrameObj.DeleteLoadStrain

## VB6 Procedure

Function DeleteLoadStrain(ByVal Name As String, ByVal LoadPat As String, ByVal DOF As Long, Optional ByVal ItemType As eItemType = Object) As Long

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

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignments are deleted for the frame object specified by the Name item.

If this item is Group, the load assignments are deleted for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignments are deleted for all selected frame objects, and the Name item is ignored.

## Remarks

This function deletes the strain load assignments to the specified frame objects, for the specified load pattern, for the specified degree of freedom.

The function returns zero if the load assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteFrameStrainLoad()

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

      ret = SapModel.FrameObj.SetLoadStrain("2", "DEAD", 1, 0.001)



   'delete frame strain load

      ret = SapModel.FrameObj.DeleteLoadStrain("1", "DEAD", 1)



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

[SetLoadStrain](SetLoadStrain_%7BFrame_Object%7D.md)
