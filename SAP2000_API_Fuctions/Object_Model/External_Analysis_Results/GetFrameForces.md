# GetFrameForces

## Syntax

SapObject.SapModel.ExternalAnalysisResults.GetFrameForce

## VB6 Procedure

Function GetFrameForce(ByVal Name As String, ByVal InitialCase As String, \_

                                       ByVal NumStep As Long, ByRef NumberStations As Long, \_

                                       ByRef P() As Double, ByRef V2() As Double, \_

                                       ByRef V3() As Double, ByRef T() As Double, \_

                                       ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

Name

The name of an existing frame object.

InitialCase

The name of an existing external results load case for which external results relevant to the object may have been previously provided.

NumStep

The zero based index of a load case step: 0 for the first step, 1 for the second, and so on.

NumberStations

The number of frame stations on the object at which external result forces are reported.

P, V2, V3

One dimensional arrays that include the axial force, shear force in the local 2 direction, and shear force in the local 3 direction, respectively, for each frame station. [F]

T, M2, M3

One dimensional arrays that include the torsion, moment about the local 2axis, and moment about the local 3-axis, respectively, for each frame station. [FL]

## Remarks

This function reports the external result forces for a given step of a given external results load case on a given object.

The function returns zero if the external result forces are successfully recovered, otherwise it returns a nonzero value.

## VBA Example

Sub GetFrameExternalResultForces ()

     

'dimension variables

        Dim SapObject As Sap2000v16.SapObject

        Dim SapModel As cSapModel

        Dim ret As Long

        Dim ObjSta() As Double = New Double() {0.0, 124.0}

        Dim LoadCase() As String = New String() {"EXTERNAL"}   

        Dim P() As Double = New Double() {5.0, 5.0}

        Dim V2() As Double = New Double() {-5.0, 5.0}

        Dim V3() As Double = Nothing

        Dim T() As Double = Nothing

        Dim M2() As Double = Nothing

        Dim M3() As Double = New Double() {100.0, 100.0}

        Dim numberStations As Long

        'create Sap2000 object

        SapObject = New Sap2000v16.SapObject

        'start Sap2000 application

        SapObject.ApplicationStart()

        'create SapModel object

        SapModel = SapObject.SapModel

        'initialize model

        ret = SapModel.InitializeNewModel()

        'create model from template

        ret = SapModel.File.New2DFrame(e2DFrameType.PortalFrame, 3, 124, 3, 200)

        'create load case for external results

        ret = SapModel.LoadCases.ExternalResults.SetCase("EXTERNAL")

        ret = SapModel.LoadCases.ExternalResults.SetNumberSteps("EXTERNAL", 1)

        'set cases and stations for frame external results

        ret = SapModel.ExternalAnalysisResults.PresetFrameCases("1", 1, LoadCase)

        ret = SapModel.ExternalAnalysisResults.SetFrameStations("1", ObjSta)

        'set frame external result forces at case first step

        ret = SapModel.ExternalAnalysisResults.SetFrameForce("1", "EXTERNAL", 0, P, V2, V3, T, M2, M3)

ret = SapModel.ExternalAnalysisResults.GetFrameForce("1", "EXTERNAL", 0, numberStations, P, V2, V3, T, M2, M3)

        'close Sap2000

 SapObject.ApplicationExit(False)

 SapModel = Nothing

       SapObject = Nothing

    End Sub

## Release Notes

Initial release in version 16.00.
