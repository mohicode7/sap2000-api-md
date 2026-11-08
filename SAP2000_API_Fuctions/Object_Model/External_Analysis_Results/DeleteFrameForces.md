# DeleteFrameForces

## Syntax

SapObject.SapModel.ExternalAnalysisResults.DeleteFrameForces

## VB6 Procedure

Function DeleteFrameForces (Name As String) As Long

## Parameters

Name

The name of an existing frame object.

## Remarks

This function deletes all the external results previously provided for a given frame object. This function returns zero if the external results are successfully deleted, otherwise it returns a nonzero value.

## VBA Example

Sub DeleteFrameExternalResults ()

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

        ret = SapModel.ExternalAnalysisResults.DeleteFrameForces("1")

        'close Sap2000

        SapObject.ApplicationExit(False)

        SapModel = Nothing

        SapObject = Nothing

   

End Sub

## Release Notes

Initial release in version 16.00.
