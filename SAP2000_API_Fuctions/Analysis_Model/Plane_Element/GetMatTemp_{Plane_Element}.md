# GetMatTemp

## Syntax

SapObject.SapModel.PlaneElm.GetMatTemp

## VB6 Procedure

Function GetMatTemp(ByVal Name As String, ByRef Temp As Double, ByRef PatternName As String) As Long

## Parameters

Name

The name of an existing plane element.

Temp

This is the material temperature value assigned to the plane element. [T]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the material temperature for the plane element is uniform over the element at the value specified by Temp.

If PatternName is the name of a defined joint pattern, the material temperature for the plane element may vary. The material temperature at each corner point around the plane element perimeter is equal to the specified temperature multiplied by the pattern value at the associated point element. The material temperature at other points in the plane element is calculated by interpolation from the corner points.

## Remarks

This function retrieves the material temperature assignments to plane elements.

The function returns zero if the material temperature assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPlaneElementMatTemp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Temp As Double

      Dim PatternName As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 3-001-incomp.sdb")



   'assign material temperature

      ret = SapModel.AreaObj.SetMatTemp("ALL", 50, , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get material temperature for plane element

      ret = SapModel.PlaneElm.GetMatTemp("3", Temp, PatternName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
