# SetReleases

## Syntax

SapObject.SapModel.FrameObj.SetReleases

## VB6 Procedure

Function SetReleases(ByVal Name As String, ByRef ii() As Boolean, ByRef jj() As Boolean, ByRef StartValue() As Double, ByRef EndValue() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

ii, jj

These are arrays of six booleans indicating the I-End and J-End releases for the frame object.

ii(0) and jj(0) = U1 release

ii(1) and jj(1) = U2 release

ii(2) and jj(2) = U3 release

ii(3) and jj(3) = R1 release

ii(4) and jj(4) = R2 release

ii(5) and jj(5) = R3 release

StartValue, EndValue

These are arrays of six values indicating the I-End and J-End partial fixity springs for the frame object.

StartValue(0) and EndValue(0) = U1 partial fixity [F/L]

StartValue(1) and EndValue(1) = U2 partial fixity [F/L]

StartValue(2) and EndValue(2) = U3 partial fixity [F/L]

StartValue(3) and EndValue(3) = R1 partial fixity [FL/rad]

StartValue(4) and EndValue(4) = R2 partial fixity [FL/rad]

StartValue(5) and EndValue(5) = R3 partial fixity [FL/rad]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function makes end release and partial fixity assignments to frame objects.

The function returns zero if the assignments are successfully retrieved, otherwise it returns a nonzero value.

Partial fixity assignments are made to degrees of freedom that have been released only.

Some release assignments would cause instability in the model. An error is returned if this type of assignment is made. Unstable release assignments include the following:

U1 released at both ends

U2 released at both ends

U3 released at both ends

R1 released at both ends

R2 released at both ends and U3 at either end

R3 released at both ends and U2 at either end

## VBA Example



```
Sub SetFrameEndReleases()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ii() As Boolean

      Dim jj() As Boolean

      Dim StartValue() As Double

      Dim EndValue() As Double



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



   'assign end releases

      ReDim ii(5)

      ReDim jj(5)

      ReDim StartValue(5)

      ReDim EndValue(5)

      ii(5) = True

      jj(5) = True

      ret = SapModel.FrameObj.SetReleases("13", ii, jj, StartValue, EndValue)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetReleases](GetReleases_%7BFrame_Object%7D.md)
