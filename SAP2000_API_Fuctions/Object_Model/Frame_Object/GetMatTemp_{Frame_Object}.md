# GetMatTemp

## Syntax

SapObject.SapModel.FrameObj.GetMatTemp

## VB6 Procedure

Function GetMatTemp(ByVal Name As String, ByRef Temp As Double, ByRef PatternName As String) As Long

## Parameters

Name

The name of an existing frame object.

Temp

This is the material temperature value assigned to the frame object. [T]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the material temperature for the frame object is uniform along the object at the value specified by Temp.

If PatternName is the name of a defined joint pattern, the material temperature for the frame object may vary from one end to the other. The material temperature at each end of the object is equal to the specified temperature multiplied by the pattern value at the joint at the end of the frame object.

## Remarks

This function retrieves the material temperature assignments to frame objects.

The function returns zero if the material temperature assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameMatTemp()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign material temperature

      ret = SapModel.FrameObj.SetMatTemp("ALL", 50, , Group)



   'get material temperature

      ret = SapModel.FrameObj.GetMatTemp("3", Temp, PatternName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetMatTemp](SetMatTemp_%7BFrame_Object%7D.md)
