# GetReleases

## Syntax

SapObject.SapModel.LineElm.GetReleases

## VB6 Procedure

Function GetReleases(ByVal Name As String, ByRef ii() As Boolean, ByRef jj() As Boolean, ByRef StartValue() As Double, ByRef EndValue() As Double) As Long

## Parameters

Name

The name of an existing line element.

ii, jj

These are arrays of six booleans indicating the I-End and J-End releases for the line element.

ii(0) and jj(0) = U1 release

ii(1) and jj(1) = U2 release

ii(2) and jj(2) = U3 release

ii(3) and jj(3) = R1 release

ii(4) and jj(4) = R2 release

ii(5) and jj(5) = R3 release

StartValue, EndValue

These are arrays of six values indicating the I-End and J-End partial fixity springs for the line element.

StartValue(0) and EndValue(0) = U1 partial fixity [F/L]

StartValue(1) and EndValue(1) = U2 partial fixity [F/L]

StartValue(2) and EndValue(2) = U3 partial fixity [F/L]

StartValue(3) and EndValue(3) = R1 partial fixity [FL/rad]

StartValue(4) and EndValue(4) = R2 partial fixity [FL/rad]

StartValue(5) and EndValue(5) = R3 partial fixity [FL/rad]

## Remarks

This function retrieves the line element end release and partial fixity assignments.

The function returns zero if the assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmEndReleases()

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



   'assign end releases to frame object

      ReDim ii(5)

      ReDim jj(5)

      ReDim StartValue(5)

      ReDim EndValue(5)

      ii(5) = True

      jj(5) = True

      ret = SapModel.FrameObj.SetReleases("13", ii, jj, StartValue, EndValue)



   'assign frame object automesh options

      ret = SapModel.FrameObj.SetAutoMesh("13", True, False, False, 2, 0)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get end releases for line element

      ReDim ii(5)

      ReDim jj(5)

      ReDim StartValue(5)

      ReDim EndValue(5)

      ret = SapModel.LineElm.GetReleases("13-1", ii, jj, StartValue, EndValue)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
