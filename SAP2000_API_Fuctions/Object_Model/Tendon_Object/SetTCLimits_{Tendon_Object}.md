# SetTCLimits

## Syntax

SapObject.SapModel.TendonObj.SetTCLimits

## VB6 Procedure

Function SetTCLimits(ByVal Name As String, ByVal LimitCompressionExists As Boolean, ByVal LimitCompression As Double, ByVal LimitTensionExists As Boolean, ByVal LimitTension As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing tendon object or group, depending on the value of the ItemType item.

LimitCompressionExists

This item is True if a compression force limit exists for the tendon object.

LimitCompression

The compression force limit for the tendon object. [F]

LimitTensionExists

This item is True if a tension force limit exists for the tendon object.

LimitTension

The tension force limit for the tendon object. [F]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the tendon object specified by the Name item.

If this item is Group, the assignment is made to all tendon objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected tendon objects and the Name item is ignored.

## Remarks

This function makes tension/compression force limit assignments to tendon objects.

The function returns zero if the assignments are successfully applied, otherwise it returns a nonzero value.

Note that the tension and compression limits are used in nonlinear analyses only.

## VBA Example



```
Sub AssignTendonTCLimits()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'assign tension/compression limits

      ret = SapModel.TendonObj.SetTCLimits(Name, True, 0, True, 100)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetTCLimits](GetTCLimits_%7BTendon_Object%7D.md)
