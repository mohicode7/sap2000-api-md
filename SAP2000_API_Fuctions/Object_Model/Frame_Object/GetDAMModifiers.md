# GetDAMModifiers

## Syntax

SapObject.SapModel.FrameObj.GetDAMModifiers

## VB6 Procedure

Function GetDAMModifiers(ByVal Name As String, ByRef EAModifier As Double, ByRef EIModifier As Double) As Long

## Parameters

Name

The name of an existing frame section whose design type is Steel Frame design.

EAModifier

The modification factor for axial stiffness if the Direct Analysis method is used.

EIModifier

The modification factor for flexural stiffness if the Direct Analysis method is used.

## **Remarks**

This function gets the modification factors for axial and flexural stiffness for a frame object if the Direct Analysis method is used.

The function returns zero if the factors are successfully retrieved, otherwise it returns a nonzero value. The function will return nonzero the modification factors are not available for the frame object.

## VBA Example

Sub GetDAMModifiers()

'dimension variables

Dim SapObject As Sap2000v16.SapObject

Dim SapModel As cSapModel

Dim ret As Long

Dim EAModifier As Double

Dim EIModifier As Double

'create Sap2000 object

Set SapObject = New Sap2000v16.SapObject

'start Sap2000 application

SapObject.ApplicationStart

'create SapModel object

Set SapModel = SapObject.SapModel

'initialize model

ret = SapModel.InitializeNewModel

'create model from template

ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)

'set the steel frame design code

ret =SapModel.DesignSteel.SetCode("AISC360-05/IBC2006")

'save model

ret = SapModel.File.Save("C:\SapAPI\x.sdb")

'run analysis

ret =SapModel.Analyze.RunAnalysis

'run design

ret =SapModel.DesignSteel.StartDesign

'get direct analysis method modification factors

ret = SapModel.FrameObj.GetDAMModifiers("1", EAModifier, EIModifier)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also
