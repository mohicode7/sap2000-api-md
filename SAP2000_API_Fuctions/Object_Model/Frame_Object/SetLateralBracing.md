# SetLateralBracing

## Syntax

SapObject.SapModel.FrameObj.SetLateralBracing

## VB6 Procedure

Function SetLateralBracing(ByVal Name As String, ByVal MyType As Long, ByVal Loc As Long, ByVal MyDist1 As Double, ByVal MyDist2 As Double, Optional ByVal RelDist As Boolean = True, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

MyType

This is 1 or 2, indicating the bracing type assigned.

1 = Point bracing

2 = Uniform bracing

Loc

This is 1, 2 or 3, indicating the bracing location.

1 = Top

2 = Bottom

3 = All (top and bottom)

MyDist1

When MyType = 1 this is the location of the point bracing.

When MyType = 2 this is the location of the start of the uniform bracing. [L] when RelDist = False

MyDist2

This item is not used when MyType = 1.

When MyType = 2 this is the location of the end of the uniform bracing. [L] when RelDist = False

RelDist

If this item is True, MyDist1 and MyDist2 are relative distances; otherwise they are actual distances.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns a lateral bracing location to frame objects.

The function returns zero if the location is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameLateralBracing()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'assign frame lateral bracing

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 0, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 0.25, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 1, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 2, 1, 0.5, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetLateralBracing](GetLateralBracing.md)

[DeleteLateralBracing](DeleteLateralBracing.md)
