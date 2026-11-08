# GetPatternValue

## Syntax

SapObject.SapModel.PointObj.GetPatternValue

## VB6 Procedure

Function GetPatternValue(ByVal Name As String, ByVal PatternName As String, ByRef Value As Double) As Long

## Parameters

Name

The name of an existing point object.

PatternName

The name of a defined joint pattern.

Value

The value that the specified point object has for the specified joint pattern.

## Remarks

This function retrieves the joint pattern value for a specific point object and joint pattern.

The function returns zero if the value is successfully retrieved, otherwise it returns a nonzero value.

Joint pattern values are unitless.

## VBA Example



```
Sub GetJointPatternData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Double



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



   'add joint pattern assignment

      ret = SapModel.PointObj.SetPatternByXYZ("ALL", "Default", 0, 0, 10, 0, Group)



   'get joint pattern assignment

      ret = SapModel.PointObj.GetPatternValue("3", "Default", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetPatternByPressure](SetPatternByPressure.md)

[SetPatternByXYZ](SetPatternByXYZ.md)

[DeletePatternValue](DeletePatternValue.md)
