# GetSP20133302016

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetSP20133302016

## VB6 Procedure

Function GetSP20133302016 (ByVal Name As String, ByRef LoadingType As Long, ByRef ExposureFrom As Long, ByRef DirAngle As Double, ByRef Cp As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef StructureType As Integer, ByRef LogDecrement As Integer, ByRef BuildingWidth As Double, ByRef BuildingDepth As Double, ByRef Zeq As Double, ByRef WindDistrict As Long, ByRef WindPressure As Double, ByRef TerrainType As Long, ByRef LimitFreq As Double, ByRef ModalCase As String, ByRef FirstValMode As Long, ByRef UserExposure As Boolean) As Long

## Parameters

Name

The name of an existing Wind-type load pattern with a SP 20.13330.2016 auto wind assignment.

LoadingType

This indicates whether the loads should be calculated as static or dynamic.

1 = Static

2 = Dynamic

ExposureFrom

This is 1 or 2, indicating the source of the wind exposure. This item applies only when LoadingType = 1.

1 = From extents of rigid diaphragms

2 = From area objects

DirAngle

The direction angle for the wind load. This item does not apply when ExposureFrom = 2.

Cp

The wind coefficient, Cp. This item applies only when LoadingType = 1 and ExposureFrom = 1.

UserZ

This item is True if the top and bottom elevations of the wind load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto wind loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto wind loads are applied. [L]

StructureType

This indicates the structure type for calculating the wind pressure.

1 = Building

2 = Other

LogDecrement

This specifies a value for the logarithmic decrement. This item applies only when LoadingType = 2.

1 = 0.15

2 = 0.30

BuildingWidth

This is the building width, normal to the wind direction. [L]

BuildingDepth

This is the building depth, in the direction of the wind. This item applies only when LoadingType = 2. [L]

Zeq

This is the equivalent height, Zeq, of the structure. This item applies only when LoadingType = 2. [L]

WindDistrict

This is the wind district to consider.

0 = Ia

1 = I

2 = II

3 = III

4 = IV

5 = V

6 = VI

7 = VII

8 = User Defined

WindPressure

This is the wind pressure. This item applies only when WindDistrict = 8. [F/L2]

TerrainType

This is the terrain type being considered.

1 = A

2 = B

3 = C

LimitFreq

The limit frequency. This item applies only when LoadingType = 2 and WindDistrict = 8.

ModalCase

The name of an existing modal load case to be used in determining dynamic wind loads. This item applies only when LoadingType = 2.

FirstValMode

The first valuable mode number from the specified modal case. This item applies only when LoadingType = 2.

UserExposure

If this item is True, the wind exposure widths are provided by the user. If it is False, the wind exposure widths are calculated by the program from the extents of the diaphragms. This item applies only when LoadingType = 1.

## Remarks

This function retrieves auto wind loading parameters for SP 20.13330.2016.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindSP20133302016()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
 
      Dim LoadingType As Long
      Dim ExposureFrom As Long
      Dim DirAngle As Double
      Dim Cp As Double
      Dim UserZ As Boolean
      Dim TopZ As Double
      Dim BottomZ As Double
      Dim StructureType As Long
      Dim LogDecrement As Long
      Dim BuildingWidth As Double
      Dim Building Depth As Double
      Dim Zeq As Double
      Dim WindDistrict As Long
      Dim WindPressure As Double
      Dim TerrainType As Long
      Dim LimitFreq As Double
      Dim ModalCase As String
      Dim FirstValMode As Long
      Dim UserExposure As Boolean
 
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
 
   'assign SP 20.13330.2016 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetSP20133302016("WIND", 1, 1, 0, 0.8, False, 0, 0, 1, 1, 10, 10, 10, 3, 0.1, 2, 10, “MODAL”, 1)
   'get SP 20.13330.2016 parameters
      ret = SapModel.LoadPatterns.AutoWind.GetSP20133302016("WIND", LoadingType, ExposureFrom, DirAngle, Cp, UserZ, TopZ, BottomZ, StructureType, LogDecrement, BuildingWidth, BuildingDepth, Zeq, WindDistrict, WindPressure, TerrainType, LimitFreq, ModalCase, FirstValMode, UserExposure)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
 
```



## Release Notes

Initial release in v21.0.0.

## See Also

[SetSP20133302016](SetSP20133302016_%7BWind_Load%7D.md)
