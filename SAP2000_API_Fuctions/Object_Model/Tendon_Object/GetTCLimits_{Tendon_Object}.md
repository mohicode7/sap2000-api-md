# GetTCLimits

## Syntax

SapObject.SapModel.TendonObj.GetTCLimits

## VB6 Procedure

Function GetTCLimits(ByVal Name As String, ByRef LimitCompressionExists As Boolean, ByRef LimitCompression As Double, ByRef LimitTensionExists As Boolean, ByRef LimitTension As Double) As Long

## Parameters

Name

The name of an existing tendon object.

LimitCompressionExists

This item is True if a compression force limit exists for the tendon object.

LimitCompression

The compression force limit for the tendon object. [F]

LimitTensionExists

This item is True if a tension force limit exists for the tendon object.

LimitTension

The tension force limit for the tendon object. [F]

## Remarks

This function retrieves the tension/compression force limit assignments to tendon objects.

The function returns zero if the assignments are successfully retrieved, otherwise it returns a nonzero value.

Note that the tension and compression limits are used in nonlinear analyses only.

## VBA Example



```
Sub GetTendonTCLimits()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim LimitCompressionExists As Boolean

      Dim LimitCompression As Double

      Dim LimitTensionExists As Boolean

      Dim LimitTension As Double



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



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'assign tension/compression limits

      ret = SapModel.TendonObj.SetTCLimits(Name, True, 0, True, 100)



   'get tension/compression limits

      ret = SapModel.TendonObj.GetTCLimits(Name, LimitCompressionExists, LimitCompression, LimitTensionExists, LimitTension)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetTCLimits](SetTCLimits_%7BTendon_Object%7D.md)
