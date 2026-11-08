# PresetFrameCases

## Syntax

SapObject.SapModel.ExternalAnalysisResults.PresetFrameCases

## VB6 Procedure

Function PresetFrameCases(Name As String, ByVal Count As Long, ByRef casname() As String) As Long

## Parameters

Name

The name of an existing frame object.

Count

The length of the subsequent array.

casename

An array listing the names of previously defined external result load cases for which user-supplied external analysis results are available for the frame object.

## Remarks

Calling this function is optional, but it can speed up subsequent assignment of external analysis results which are available for more than one load case.

The first time this function is called for a particular frame object, it sets the list of names of external result load cases for which results relevant to the object are available. Subsequent calls to this function for the same object reset the results for load cases already in the list, and add load cases not already in.

This function returns zero if the stations are successfully set, otherwise it returns a nonzero value.

## VBA Example

Sub PresetFrameExternalCases()

    'dimension variables

         Dim SapObject as cOAPI

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
