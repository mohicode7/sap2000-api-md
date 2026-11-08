# GetModifiers

## Syntax

SapObject.SapModel.LineElm.GetModifiers

## VB6 Procedure

Function GetModifiers(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing line element.

Value

This is an array of eight unitless modifiers.

Value(0) = Cross sectional area modifier

Value(1) = Shear area in local 2 direction modifier

Value(2) = Shear area in local 3 direction modifier

Value(3) = Torsional constant modifier

Value(4) = Moment of inertia about local 2 axis modifier

Value(5) = Moment of inertia about local 3 axis modifier

Value(6) = Mass modifier

Value(7) = Weight modifier

## Remarks

This function retrieves the section  modifier assignment for line elements. The default value for all modifiers is one.

The function returns zero if the modifier assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmModifiers()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Value() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign modifiers to frame objects

      ReDim Value(7)

      For i = 0 To 7

         Value(i) = 1

      Next i

      Value(5) = 100

      ret = SapModel.FrameObj.SetModifiers("3", Value)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get modifiers for line element

      ReDim Value(7)

      ret = SapModel.LineElm.GetModifiers("3-1", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
