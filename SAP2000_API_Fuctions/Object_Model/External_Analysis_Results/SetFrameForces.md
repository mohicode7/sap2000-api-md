# SetFrameForce

## Syntax

SapObject.SapModel.ExternalAnalysisResults.SetFrameForce

## VB6 Procedure

Function SetFrameForce(ByVal Name As String, ByVal InitialCase As String, ByVal StepNum As Long, ByRef P() As Double, ByRef V2() As Double, ByRef V3() As Double, ByRef T() As Double, ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

Name

The name of an existing frame object.

InitialCase

The name of an existing external results load case for which user-supplied external analysis results relevant to the object are available.

StepNum

The zero based index of a load case step: 0 for the first step, 1 for the second, and so on.

P, V2, V3

These are one dimensional arrays that include the axial force, shear force in the local 2 direction, and shear force in the local 3 direction, respectively, for each frame station.[F] These arrays are each expected to contain a number of values equal to the  previously declared number of stations at which external results are available for the object – see SapObject.SapModel.ExternalAnalysisResults.SetFrameStations() for that declaration.  

If any of these arrays is empty, the function substitutes zero values for the missing values.

T, M2, M3

These are one dimensional arrays that include the torsion, moment about the local 2axis, and moment about the local 3-axis, respectively, for each frame station. [FL] Again, these arrays are each expected to contain a number of values equal to the  previously declared number of stations at which external results are provided for the object. If any of these arrays is empty, the function substitutes zero values for the missing values.

## Remarks

This function sets the external result forces for a given step of a given external results load case on a frame object.

The function returns zero if the forces are successfully set, otherwise it returns a nonzero value.

## VBA Example

Sub SetFrameExternalResultForces ()

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

        'create Sap2000 object

        SapObject = New Sap2000v16.SapObject

        'start Sap2000 application

        SapObject.ApplicationStart()

        'create SapModel object

        SapModel = SapObject.SapModel

        'initialize model

        ret = SapModel.InitializeNewModel

        'create model from template

        ret = SapModel.File.New2DFrame(e2DFrameType.PortalFrame, 3, 124, 3, 200)

        'create load case for external results

        ret = SapModel.LoadCases.ExternalResults.SetCase("EXTERNAL")

       

        'set cases and stations for frame external results

        ret = SapModel.ExternalAnalysisResults.PresetFrameCases("1", 1, LoadCase)

        ret = SapModel.ExternalAnalysisResults.SetFrameStations("1", ObjSta)

        'set frame external result forces at case first step

        ret = SapModel.ExternalAnalysisResults.SetFrameForce("1", "EXTERNAL", 0, P, V2, V3, T, M2, M3)

        'close Sap2000

        SapObject.ApplicationExit(False)

        SapModel = Nothing

        SapObject = Nothing

    End Sub

## Release Notes

Initial release in version 16.00.
