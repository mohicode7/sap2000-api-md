# GetNBCC95

## Syntax

SapObject.SapModel.Func.FuncRS.GetNBCC95

## VB6 Procedure

Function GetNBCC95(ByVal Name As String, ByRef NBCC95ZVR As Double, ByRef NBCC95ZA As Long, ByRef NBCC95ZV As Long, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a NBCC95 response spectrum function.

NBCC95ZVR

The zonal velocity ratio.

NBCC95ZA

The acceleration-related seismic zone.

NBCC95ZV

The velocity-related seismic zone.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a NBCC95 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncNBCC95()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NBCC95ZVR As Double

      Dim NBCC95ZA As Long

      Dim NBCC95ZV As Long

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



   'add NBCC95 RS function

      ret = SapModel.Func.FuncRS.SetNBCC95("RS-1", 0.3, 5, 5, 0.04)



   'get NBCC95 RS function

      ret = SapModel.Func.FuncRS.GetNBCC95("RS-1", NBCC95ZVR, NBCC95ZA, NBCC95ZV, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetNBCC95](SetNBCC95_%7BRS%7D.md)
