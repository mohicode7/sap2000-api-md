# SetReleases

## Syntax

SapObject.SapModel.NamedAssign.ReleaseFrame.SetReleases

## VB6 Procedure

Function SetReleases(ByVal Name As String, ByRef ii() As Boolean, ByRef jj() As Boolean, ByRef StartValue() As Double, ByRef EndValue() As Double) As Long

## Parameters

Name

The name of a new or existing frame end release.

ii, jj

These are arrays of six booleans indicating the I-End and J-End releases.

ii(0) and jj(0) = U1 release

ii(1) and jj(1) = U2 release

ii(2) and jj(2) = U3 release

ii(3) and jj(3) = R1 release

ii(4) and jj(4) = R2 release

ii(5) and jj(5) = R3 release

StartValue, EndValue

These are arrays of six values indicating the I-End and J-End partial fixity springs.

StartValue(0) and EndValue(0) = U1 partial fixity [F/L]

StartValue(1) and EndValue(1) = U2 partial fixity [F/L]

StartValue(2) and EndValue(2) = U3 partial fixity [F/L]

StartValue(3) and EndValue(3) = R1 partial fixity [FL/rad]

StartValue(4) and EndValue(4) = R2 partial fixity [FL/rad]

StartValue(5) and EndValue(5) = R3 partial fixity [FL/rad]

## Remarks

This function defines a named frame end release.

The function returns zero if the release is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub DefineFrameEndReleases()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ii() As Boolean

      Dim jj() As Boolean

      Dim StartValue() As Double

      Dim EndValue() As Double



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



   'define releases

      ReDim ii(5)

      ReDim jj(5)

      ReDim StartValue(5)

      ReDim EndValue(5)

      ii(0) = True

      StartValue(0) = 10

      ret = SapModel.NamedAssign.ReleaseFrame.SetReleases("FREL1", ii, jj, StartValue, EndValue)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetReleases](GetReleases_%7BFrame_Releases%7D.md)
