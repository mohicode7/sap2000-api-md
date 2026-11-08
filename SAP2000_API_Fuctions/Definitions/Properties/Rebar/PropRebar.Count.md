# Count

## Syntax

SapObject.SapModel.PropRebar.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the total number of defined rebar properties in the model.

## VBA Example



```
Sub CountRebarProps()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim Count As Long



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



       'return number of defined rebar properties

          Count = SapModel.PropRebar.Count



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## **Release Notes**

Initial release in v21.2.0.

## See Also
