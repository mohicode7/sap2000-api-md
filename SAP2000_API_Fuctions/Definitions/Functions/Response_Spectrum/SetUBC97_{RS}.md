# SetUBC97

## Syntax

SapObject.SapModel.Func.FuncRS.SetUBC97

## VB6 Procedure

Function SetUBC97(ByVal Name As String, ByVal UBC97Ca As Double, ByVal UBC97Cv As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

UBC97Ca

The seismic coefficient, Ca.

UBC97Cv

The seismic coefficient, Cv.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a UBC97 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncUBC97()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetUBC97](GetUBC97_%7BRS%7D.md)
