# GetExposure

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetExposure

## VB6 Procedure

Function GetExposure(ByVal Name As String, ByRef Num As Long, ByRef Diaph() As String, ByRef x() As Double, ByRef y() As Double, ByRef MyWidth() As Double, ByRef Height() As Double) As Long

## Parameters

Name

The name of an existing Wind-type load pattern that has an auto wind load assigned.

Num

The number of diaphragms at which exposure data is reported.

Diaph

This is an array that includes the names of the diaphragms that have eccentricity overrides.

x

This is an array that includes the global X-coordinate of the point where the wind force load is applied to the diaphragm. [L]

y

This is an array that includes the global Y-coordinate of the point where the wind force load is applied to the diaphragm. [L]

MyWidth

This is an array that includes the exposure width for the wind load applied to the specified diaphragm. [L]

Height

This is an array that includes the exposure height for the wind load applied to the specified diaphragm. [L]

## Remarks

This function retrieves exposure parameters for auto wind loads determined from extents of rigid diaphragms. This function does not apply for User-type auto wind loads.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindExposure()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Num As Long

      Dim Diaph() As String

      Dim x() As Double

      Dim y() As Double

      Dim MyWidth() As Double

      Dim Height() As Double



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

      ret = SapModel.LoadPatterns.AutoWind.SetExposure("WIND", "Diaph2", 0, 0, 900, 125)



   'get exposure data

      ret = SapModel.LoadPatterns.AutoWind.GetExposure("WIND", Num, Diaph, x, y, MyWidth, Height)



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

[SetExposure](SetExposure.md)

[GetSpecialRigidDiaphragmList](../Definitions/Constraints/GetSpecialRigidDiaphragmList.md)
