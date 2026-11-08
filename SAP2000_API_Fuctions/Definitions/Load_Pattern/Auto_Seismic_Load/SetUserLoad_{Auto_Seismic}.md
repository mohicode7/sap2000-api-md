# SetUserLoad

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetUserLoad

## VB6 Procedure

Function SetUserLoad(ByVal Name As String, ByVal MyType As Long, Optional ByVal Eccen As Double = 0.05) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

MyType

This is either 1 or 2, indicating the application point type for the user load.

1 = User specified application point

2 = At center of mass with optional additional eccentricity

Eccen

The eccentricity ratio that applies to all diaphragms. This item is only applicable when MyType = 2.

## Remarks

This function sets the auto seismic load type to User Load. User load values are assigned using the SetUserLoadValue function.

The function returns zero if the load type is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicUserLoad()

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

[SetUserLoadValue](SetUserLoadValue.md)

[GetUserLoad](GetUserLoad_%7BAuto_Seismic%7D.md)
