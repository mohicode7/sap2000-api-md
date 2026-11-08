# SetExposure

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.SetExposure

## VB6 Procedure

Function SetExposure(ByVal Name As String, ByVal Diaph As String, ByVal x As Double, ByVal y As Double, ByVal Width As Double, ByVal Height As Double) As Long

## Parameters

Name

The name of an existing Wind-type load pattern that has an auto wind load assigned.

Diaph

The name of an existing special rigid diaphragm constraint, that is, a diaphragm constraint with the following features:

1. The constraint type is CONSTRAINT\_DIAPHRAGM = 2.

2. The constraint coordinate system is Global.

3. The constraint axis is Z.

x

The global X-coordinate of the point where the wind force is applied. [L]

y

The global Y-coordinate of the point where the wind force is applied. [L]

Width

The exposure width for the wind load applied to the specified diaphragm. [L]

Height

The exposure height for the wind load applied to the specified diaphragm. [L]

## Remarks

This function assigns exposure parameters for auto wind loads determined from extents of rigid diaphragms. This function does not apply for User-type auto wind loads.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value. The function returns an error if the auto wind load is not specified to have user exposure parameters.

## VBA Example



```
Sub AssignWindUserExposure()

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

      ret = SapModel.LoadPatterns.Add("WIND", LTYPE_WIND)



   'assign ASCE788 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetASCE788("WIND", 1, 0, 0.8, 0.5, False, 0, 0, 80, 3, 1, 0.85, True)



   'assign user exposure data

      ret = SapModel.LoadPatterns.AutoWind.SetExposure("WIND", "Diaph2", 0, 0, 900, 124)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetExposure](GetExposure.md)

[GetSpecialRigidDiaphragmList](../Definitions/Constraints/GetSpecialRigidDiaphragmList.md)
