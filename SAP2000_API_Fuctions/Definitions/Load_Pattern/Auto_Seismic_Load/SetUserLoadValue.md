# SetUserLoadValue

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetUserLoadValue

## VB6 Procedure

Function SetUserLoadValue(ByVal Name As String, ByVal Diaph As String, ByVal Fx As Double, ByVal Fy As Double, ByVal Mz As Double, Optional ByVal x As Double = 0, Optional ByVal y As Double = 0) As Long

## Parameters

Name

The name of an existing Quake-type load pattern that has been assigned a User Load auto seismic type.

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

The global X-coordinate of the point where the seismic force is applied. [L]

This item is applicable only when the auto seismic load is specified to have a user specified application point (see the SetUserLoad function).

y

The global Y-coordinate of the point where the seismic force is applied. [L]

This item is applicable only when the auto seismic load is specified to have a user specified application point (see the SetUserLoad function).

## Remarks

This function assigns loading to an auto seismic load that is a User Load type. The SetUserLoad function is used to specify that an auto seismic load is a User Load type

The function returns zero if the loading is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicUserLoadValue()

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



   'set to auto seismic user load

      ret = SapModel.LoadPatterns.AutoSeismic.SetUserLoad("EQX", 1)



   'set to auto seismic user load values

      ret = SapModel.LoadPatterns.AutoSeismic.SetUserLoadValue("EQX", "Diaph1", 20, 4, 1000, 0, 0)

      ret = SapModel.LoadPatterns.AutoSeismic.SetUserLoadValue("EQX", "Diaph2", 40, 8, 2000, 0, 0)



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

[SetUserLoad](SetUserLoad_%7BAuto_Seismic%7D.md)

[GetUserLoad](GetUserLoad_%7BAuto_Seismic%7D.md)
