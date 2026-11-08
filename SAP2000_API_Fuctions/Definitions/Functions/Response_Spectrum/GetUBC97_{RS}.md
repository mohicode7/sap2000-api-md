# GetUBC97

## Syntax

SapObject.SapModel.Func.FuncRS.GetUBC97

## VB6 Procedure

Function GetUBC97(ByVal Name As String, ByRef UBC97Ca As Double, ByRef UBC97Cv As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a UBC97 response spectrum function.

UBC97Ca

The seismic coefficient, Ca.

UBC97Cv

The seismic coefficient, Cv.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a UBC97 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncUBC97()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim UBC97Ca As Double

      Dim UBC97Cv As Double

      Dim DampRatio As Double



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



   'add UBC97 RS function

      ret = SapModel.Func.FuncRS.SetUBC97("RS-1", 0.36, 0.54, 0.04)



   'get UBC97 RS function

      ret = SapModel.Func.FuncRS.GetUBC97("RS-1", UBC97Ca, UBC97Cv, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetUBC97](GetUBC97_%7BRS%7D.md)
