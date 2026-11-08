# SetNBCC95

## Syntax

SapObject.SapModel.Func.FuncRS.SetNBCC95

## VB6 Procedure

Function SetNBCC95(ByVal Name As String, ByVal NBCC95ZVR As Double, ByVal NBCC95ZA As Long, ByVal NBCC95ZV As Long, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function then that function is modified; otherwise, a new function is added.

NBCC95ZVR

The zonal velocity ratio.

NBCC95ZA

The acceleration-related seismic zone.

NBCC95ZV

The velocity-related seismic zone.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a NBCC95 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncNBCC95()

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



   'add NBCC95 RS function

      ret = SapModel.Func.FuncRS.SetNBCC95("RS-1", 0.3, 5, 5, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetNBCC95](GetNBCC95_%7BRS%7D.md)
