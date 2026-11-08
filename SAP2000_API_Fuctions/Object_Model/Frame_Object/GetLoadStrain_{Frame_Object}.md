# GetLoadStrain

## Syntax

SapObject.SapModel.FrameObj.GetLoadStrain

## VB6 Procedure

Function GetLoadStrain(ByVal Name As String, ByRef NumberItems As Long, ByRef FrameName() As String, ByRef LoadPat() As String, ByRef DOF() As Long, ByRef Val() As Double, ByRef PatternName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

NumberItems

The total number of strain loads retrieved for the specified frame objects.

FrameName

This is an array that includes the name of the frame object associated with each strain load.

LoadPat

This is an array that includes the name of the load pattern associated with each strain load.

DOF

This is an array that includes 1, 2, 3, 4, 5 or 6, indicating the degree of freedom associated with each strain load.

1 = Strain11

2 = Strain12

3 = Strain13

4 = Curvature1

5 = Curvature2

6 = Curvature3

Val

This is an array that includes the strain value. [L/L] for DOF = 1, 2 and 3 and [1/L] for DOF = 4, 5 and 6

PatternName

This is an array that includes the joint pattern name, if any, used to specify the strain load.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the frame object specified by the Name item.

If this item is Group, the assignments are retrieved for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected frame objects, and the Name item is ignored.

## Remarks

This function retrieves the strain load assignments to frame objects.

The function returns zero if the strain load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameStrainLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim FrameName() As String

      Dim LoadPat() As String

      Dim DOF() As Long

      Dim Val() As Double

      Dim PatternName() As String



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



   'get frame strain load

      ret = SapModel.FrameObj.GetLoadStrain("1", NumberItems, FrameName, LoadPat, DOF, Val, PatternName)



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

[SetLoadStrain](SetLoadStrain_%7BFrame_Object%7D.md)

[DeleteLoadStrain](DeleteLoadStrain_%7BFrame_Object%7D.md)
