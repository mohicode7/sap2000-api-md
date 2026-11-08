# SetDiaphragmEccentricityOverride

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetDiaphragmEccentricityOverride

## VB6 Procedure

Function SetDiaphragmEccentricityOverride(ByVal Name As String, ByVal Diaph As String, ByVal Eccen As Double, Optional ByVal Delete As Boolean = False) As Long

## Parameters

Name

The name of an existing Quake-type load pattern that has an auto seismic load assigned.

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

This function assigns diaphragm eccentricity overrides for auto seismic loads. This function does not apply for User Load type auto seismic loads.

The function returns zero if the overrides are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicDiaphragmEccentricityOverride()

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



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("EQX", LTYPE_QUAKE)



   'assign BOCA96 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetBOCA96("EQX", 1, 0.05, 1, 0.035, 0, False, 0, 0, 0.4, 0.4, 1.5, 8)



   'assign diaphragm eccentricity override

      ret = SapModel.LoadPatterns.AutoSeismic.SetDiaphragmEccentricityOverride("EQX", "Diaph1", 50)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Modified optional argument Delete to be ByVal in version 12.0.1.

## See Also

[GetDiaphragmEccentricityOverride](GetDiaphragmEccentricityOverride_%7BAuto_Seismic%7D.md)

[GetSpecialRigidDiaphragmList](../../Constraints/GetSpecialRigidDiaphragmList.md)
