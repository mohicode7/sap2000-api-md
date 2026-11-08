# GetUserLoad

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetUserLoad

## VB6 Procedure

Function GetUserLoad(ByVal Name As String, ByRef Num As Long, ByRef Diaph() As String, ByRef Fx() As Double, ByRef Fx() As Double, ByRef Mz() As Double, ByRef x() As Double, ByRef y() As Double) As Long

## Parameters

Name

The name of an existing Wind-type load pattern that has a User-type auto wind load assigned.

Num

The number of diaphragms at which user wind loads are reported.

Diaph

This is an array that includes the names of the diaphragms that have user wind loads.

Fx

This is an array that includes the global X direction force assigned to the specified diaphragm. [F]

Fy

This is an array that includes the global Y direction force assigned to the specified diaphragm. [F]

Mz

This is an array that includes the moment about the global Z axis assigned to the specified diaphragm. [FL]

x

This is an array that includes the global X-coordinate of the point where the wind force load is applied to the diaphragm. [L]

y

This is an array that includes the global Y-coordinate of the point where the wind force load is applied to the diaphragm. [L]

## Remarks

This function retrieves auto wind loading parameters for User-type wind loading.

The function returns zero if the parameters are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindUser()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Num As Long

      Dim Diaph() As String

      Dim Fx() As Double

      Dim Fy() As Double

      Dim Mz() As Double

      Dim x() As Double

      Dim y() As Double



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



   'get user wind load parameters

      ret = SapModel.LoadPatterns.AutoWind.GetUserLoad("WIND", Num, Diaph, Fx, Fy, Mz, x, y)



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

[SetUserLoad](SetUserLoad_%7BWind_Load%7D.md)

[GetSpecialRigidDiaphragmList](../../Constraints/GetSpecialRigidDiaphragmList.md)
