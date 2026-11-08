# Delete

## Syntax

SapObject.SapModel.PropRebar.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

**Name**

The name of an existing rebar property.

## Remarks

The function deletes a specified rebar property.

The function returns zero if the property is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified property can not be deleted, for example, if it is assigned to an existing object.

## VBA Example



```
Sub DeleteRebarProp()

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



       'set new rebar properties

          ret = SapModel.PropRebar.SetProp("R1", 1, 1.128)

          ret = SapModel.PropRebar.SetProp("R2", 1, 1.128)



       'delete rebar property

          ret = SapModel.PropRebar.Delete("R1")



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## **Release Notes**

Initial release in v21.2.0.

## **See Also**
