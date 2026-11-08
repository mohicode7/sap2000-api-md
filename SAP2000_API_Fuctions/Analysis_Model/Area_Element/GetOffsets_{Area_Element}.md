# GetOffsets

## Syntax

SapObject.SapModel.AreaElm.GetOffsets

## VB6 Procedure

Function GetOffsets(ByVal Name As String, ByRef OffsetType As Long, ByRef OffsetPattern As String, ByRef OffsetPatternSF As Double, ByRef Offset() As Double) As Long

## Parameters

Name

The name of an existing area element.

OffsetType

This is 0, 1 or 2, indicating the joint offset type.

0 = No joint offsets

1 = User defined joint offsets specified by joint pattern

2 = User defined joint offsets specified by point

OffsetPattern

This item applies only when OffsetType = 1. It is the name of the defined joint pattern that is used to calculate the joint offsets.

OffsetPatternSF

This item only applies when OffsetType = 1. It is the scale factor applied to the joint pattern when calculating the joint offsets. [L]

Offset

This item applies only when OffsetType = 2. It is an array of joint offsets for each of the points that define the area element. [L]

## Remarks

This function retrieves the joint offset assignments for area elements.

The function returns zero if the assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaElementJointOffsets()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i as long

      Dim OffsetType As Long

      Dim OffsetPattern As String

      Dim OffsetPatternSF As Double

      Dim Offset() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign joint offsets

      ReDim Offset(3)

      For i = 0 To 3

         Offset(i) = 12

      Next i

      ret = SapModel.AreaObj.SetOffsets("ALL", 2, "", 1, Offset, Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get joint offsets for area element

      ret = SapModel.AreaElm.GetOffsets("3", OffsetType, OffsetPattern, OffsetPatternSF, Offset)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
