# GetUserLoad

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetUserLoad

## VB6 Procedure

Function GetUserLoad(ByVal Name As String, ByRef MyType As Long, ByRef Eccen As Double, ByRef Num As Long, ByRef Diaph() As String, ByRef Fx() As Double, ByRef Fy() As Double, ByRef Mz() As Double, ByRef x() As Double, ByRef y() As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern which has been assigned a User Load auto seismic type.

MyType

This is either 1 or 2, indicating the application point type for the user load.

1 = User specified application point

2 = At center of mass with optional additional eccentricity

Eccen

The eccentricity ratio that applies to all diaphragms. This item is only applicable when MyType = 2.

Num

The number of diaphragms that can be loaded by the auto seismic load.

Diaph

This is an array that includes the names of the diaphragms that can be loaded by the auto seismic load.

Fx

This is an array that includes the global X direction force assigned to each diaphragm. [F]

Fy

This is an array that includes the global Y direction force assigned to each diaphragm. [F]

Mz

This is an array that includes the moment about the global Z axis for each diaphragm. [FL]

x

This is an array that includes the global X-coordinate of the point where the seismic force is applied to each diaphragm. [L]

This item is applicable only when MyType = 1.

y

This is an array that includes the global Y-coordinate of the point where the seismic force is applied to each diaphragm. [L]

This item is applicable only when MyType = 1.

## Remarks

This function retrieves auto seismic loading parameters for User Load type auto seismic loads.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicUserLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long

      Dim Eccen As Double

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

      ret = SapModel.LoadPatterns.Add("EQX", LTYPE_QUAKE)



   'set to auto seismic user load

      ret = SapModel.LoadPatterns.AutoSeismic.SetUserLoad("EQX", 1)



   'set to auto seismic user load values

      ret = SapModel.LoadPatterns.AutoSeismic.SetUserLoadValue("EQX", "Diaph1", 20, 4, 1000, 0, 0)

      ret = SapModel.LoadPatterns.AutoSeismic.SetUserLoadValue("EQX", "Diaph2", 40, 8, 2000, 0, 0)



   'get auto seismic user load parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetUserLoad("EQX", MyType, Eccen, Num, Diaph, Fx, Fy, Mz, x, y)



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

[SetUserLoadValue](SetUserLoadValue.md)
