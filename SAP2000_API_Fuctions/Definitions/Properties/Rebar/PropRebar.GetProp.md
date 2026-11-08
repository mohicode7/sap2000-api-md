# GetProp

## Syntax

SapObject.SapModel.PropRebar.GetProp

## VB6 Procedure

Function GetProp(ByVal Name As String, ByRef Area As Double, ByRef Diameter As Double) As Long

## Parameters

Name

The name of an existing rebar property.

**Area**

The cross-sectional area of the rebar. [L2]

**Diameter**

The diameter of the rebar. [L]

## Remarks

This function retrieves rebar property definition data.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRebarProperty()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim Area As Double

          Dim Diameter As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



       'set new rebar property

          ret = SapModel.PropRebar.SetProp("R1", 1, 1.128)



       'get rebar property data

          ret = SapModel.PropRebar.GetProp("R1", Area, Diameter)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in v21.2.0.

## See Also

[SetProp](PropRebar.SetProp.md)
