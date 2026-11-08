# SetUserLoad

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.SetUserLoad

## VB6 Procedure

Function SetUserLoad(ByVal Name As String, ByVal Diaph As String, ByVal Fx As Double, ByVal Fy As Double, ByVal Mz As Double, ByVal x As Double, ByVal y As Double) As Long

## Parameters

Name

The name of an existing Wind-type load pattern that has an auto wind load assigned.

Diaph

The name of an existing special rigid diaphragm constraint, that is, a diaphragm constraint with the following features:

1. The constraint type is CONSTRAINT\_DIAPHRAGM = 2.

2. The constraint coordinate system is Global.

3. The constraint axis is Z.

Fx

The global X direction force assigned to the specified diaphragm. [F]

Fy

The global Y direction force assigned to the specified diaphragm. [F]

Mz

The moment about the global Z axis assigned to the specified diaphragm. [FL]

x

The global X-coordinate of the point where the wind force is applied. [L]

y

The global Y-coordinate of the point where the wind force is applied. [L]

## Remarks

This function assigns auto wind loading parameters for User-type wind loads.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignWindUserLoad()

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



   'assign user wind load parameters

      ret = SapModel.LoadPatterns.AutoWind.SetUserLoad("WIND", "Diaph1", 100, 20, 5000, 0, 0)

      ret = SapModel.LoadPatterns.AutoWind.SetUserLoad("WIND", "Diaph2", 50, 10, 2500, 0, 0)



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

[GetUserLoad](GetUserLoad_%7BWind_Load%7D.md)

[GetSpecialRigidDiaphragmList](../../Constraints/GetSpecialRigidDiaphragmList.md)
