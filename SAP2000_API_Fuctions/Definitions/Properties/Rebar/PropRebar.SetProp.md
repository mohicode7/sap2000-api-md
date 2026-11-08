# SetProp

## Syntax

SapObject.SapModel.PropRebar.SetProp

## VB6 Procedure

Function SetProp(ByVal Name As String, ByVal Area As Double, ByVal Diameter As Double) As Long

## Parameters

**Name**

The name of an existing or new rebar property. If this is an existing property, that property is modified; otherwise, a new property is added.

**Area**

The cross-sectional area of the rebar. [L2]

**Diameter**

The diameter of the rebar. [L]

## Remarks

This function defines a rebar property.

The function returns zero if the property is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRebarProperty()

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

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



       'set new rebar property

          ret = SapModel.PropRebar.SetProp("R1", 1, 1.128)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in v21.2.0.

## See Also

[GetProp](PropRebar.GetProp.md)
