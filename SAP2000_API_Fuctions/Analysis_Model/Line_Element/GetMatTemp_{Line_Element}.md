# GetMatTemp

## Syntax

SapObject.SapModel.LineElm.GetMatTemp

## VB6 Procedure

Function GetMatTemp(ByVal Name As String, ByRef Temp As Double, ByRef PatternName As String) As Long

## Parameters

Name

The name of an existing line element.

Temp

This is the material temperature value assigned to the line element. [T]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the material temperature for the line element is uniform along the element at the value specified by Temp.

If PatternName is the name of a defined joint pattern, the material temperature for the line element may vary from one end to the other. The material temperature at each end of the element is equal to the specified temperature multiplied by the pattern value at the joint at the end of the line element.

## Remarks

This function retrieves the material temperature assignments to line elements.

The function returns zero if the material temperature assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmMatTemp()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'assign material temperature to frame objects

      ret = SapModel.FrameObj.SetMatTemp("ALL", 50, , Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get material temperature for line element

      ret = SapModel.LineElm.GetMatTemp("1-1", Temp, PatternName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
