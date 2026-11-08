# GetThickness

## Syntax

SapObject.SapModel.AreaObj.GetThickness

## VB6 Procedure

Function GetThickness(ByVal Name As String, ByRef ThicknessType As Long, ByRef ThicknessPattern As String, ByRef ThicknessPatternSF As Double, ByRef Thickness() As Double) As Long

## Parameters

Name

The name of an existing area object.

ThicknessType

This is 0, 1 or 2, indicating the thickness overwrite type.

0 = No thickness overwrites

1 = User defined thickness overwrites specified by joint pattern

2 = User defined thickness overwrites specified by point

ThicknessPattern

This item applies only when ThicknessType = 1. It is the name of the defined joint pattern that is used to calculate the thicknesses.

ThicknessPatternSF

This item applies only when ThicknessType = 1. It is the scale factor applied to the joint pattern when calculating the thicknesses. [L]

Thickness

This item applies only when ThicknessType = 2. It is an array of thicknesses at each of the points that define the area object. [L]

## Remarks

This function retrieves the thickness overwrite assignments for area objects.

The function returns zero if the assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectThicknessOverwrites()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i as long

      Dim ThicknessType As Long

      Dim ThicknessPattern As String

      Dim ThicknessPatternSF As Double

      Dim Thickness() As Double



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



   'assign thickness overwrites

      ReDim Thickness(3)

      For i = 0 To 3

         Thickness(i) = 11

      Next i

      ret = SapModel.AreaObj.SetThickness("ALL", 2, "", 1, Thickness, Group)



   'get thickness overwrites

      ret = SapModel.AreaObj.GetThickness("3", ThicknessType, ThicknessPattern, ThicknessPatternSF, Thickness)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetThickness](SetThickness.md)
