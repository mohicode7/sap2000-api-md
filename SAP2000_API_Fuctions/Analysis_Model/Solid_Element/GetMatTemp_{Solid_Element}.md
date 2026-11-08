# GetMatTemp

## Syntax

SapObject.SapModel.SolidElm.GetMatTemp

## VB6 Procedure

Function GetMatTemp(ByVal Name As String, ByRef Temp As Double, ByRef PatternName As String) As Long

## Parameters

Name

The name of an existing solid element.

Temp

This is the material temperature value assigned to the solid element. [T]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the material temperature for the solid element is uniform over the element at the value specified by Temp.

If PatternName is the name of a defined joint pattern, the material temperature for the solid element may vary. The material temperature at each corner point of the solid element is equal to the specified temperature multiplied by the pattern value at the associated point element. The material temperature at other locations in the solid element is calculated by interpolation from the corner points.

## Remarks

This function retrieves the material temperature assignments to solid elements.

The function returns zero if the material temperature assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidElementMatTemp()

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



   'create model from template

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign material temperature

      ret = SapModel.SolidObj.SetMatTemp("ALL", 50, , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get material temperature for solid element

      ret = SapModel.SolidElm.GetMatTemp("1", Temp, PatternName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
