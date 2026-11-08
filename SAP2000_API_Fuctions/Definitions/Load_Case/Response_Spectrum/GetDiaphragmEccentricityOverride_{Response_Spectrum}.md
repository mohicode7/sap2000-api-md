# GetDiaphragmEccentricityOverride

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.GetDiaphragmEccentricityOverride

## VB6 Procedure

Function GetDiaphragmEccentricityOverride(ByVal Name As String, ByRef Num As Long, ByRef Diaph() As String, ByRef Eccen() As Double) As Long

## Parameters

Name

The name of an existing response spectrum load case.

Num

The number of diaphragm eccentricity overrides for the specified load case.

Diaph

This is an array that includes the names of the diaphragms that have eccentricity overrides.

Eccen

This is an array that includes the eccentricity applied to each diaphragm. [L]

## Remarks

This function retrieves the diaphragm eccentricity overrides for a response spectrum load case.

The function returns zero if the overrides are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetResponseSpectrumDiaphragmEccentricityOverride()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Num As Long

      Dim Diaph() As String

      Dim Eccen() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New3DFrame(BeamSlab, 2, 144, 3, 336, 2, 432)



   'define diaphragm constraints

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1", Z)

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph2", Z)



   'assign points to diaphragm

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.SelectObj.PlaneXY("2")

      ret = SapModel.PointObj.SetConstraint("", "Diaph1", SelectedObjects)

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.SelectObj.PlaneXY("3")

      ret = SapModel.PointObj.SetConstraint("", "Diaph2", SelectedObjects)

      ret = SapModel.SelectObj.ClearSelection



   'add response spectrum load case

      ret = SapModel.LoadCases.ResponseSpectrum.SetCase("LCASE1")



   'assign diaphragm eccentricity override

      ret = SapModel.LoadCases.ResponseSpectrum.SetDiaphragmEccentricityOverride("LCASE1", "Diaph1", 50)



   'get diaphragm eccentricity override

      ret = SapModel.LoadCases.ResponseSpectrum.GetDiaphragmEccentricityOverride("LCASE1", Num, Diaph, Eccen)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SetDiaphragmEccentricityOverride](SetDiaphragmEccentricityOverride_%7BResponse_Spectrum%7D.md)

[GetSpecialRigidDiaphragmList](../../Constraints/GetSpecialRigidDiaphragmList.md)
