# SetDiaphragmEccentricityOverride

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.SetDiaphragmEccentricityOverride

## VB6 Procedure

Function SetDiaphragmEccentricityOverride(ByVal Name As String, ByVal Diaph As String, ByVal Eccen As Double, Optional ByVal Delete As Boolean = False) As Long

## Parameters

Name

The name of an existing response spectrum load case.

Diaph

The name of an existing special rigid diaphragm constraint, that is, a diaphragm constraint with the following features:

1. The constraint type is CONSTRAINT\_DIAPHRAGM = 2.

2. The constraint coordinate system is Global.

3. The constraint axis is Z.

Eccen

The eccentricity applied to the specified diaphragm. [L]

Delete

If this item is True, the eccentricity override for the specified diaphragm is deleted.

## Remarks

This function assigns diaphragm eccentricity overrides for response spectrum load cases.

The function returns zero if the overrides are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignResponseSpectrumDiaphragmEccentricityOverride()

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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Modified optional argument Delete to be ByVal in version 12.0.1.

## See Also

[GetDiaphragmEccentricityOverride](GetDiaphragmEccentricityOverride_%7BResponse_Spectrum%7D.md)

[GetSpecialRigidDiaphragmList](../../Constraints/GetSpecialRigidDiaphragmList.md)
