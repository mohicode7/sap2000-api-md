# GetBridgeUpdateForAnalysisFlag (Note:  Newer function available)

## Syntax

SapObject.SapModel.BridgeObj.GetBridgeUpdateForAnalysisFlag

## VB6 Procedure

Function GetBridgeUpdateForAnalysisFlag() As Boolean

## Parameters

None

## Remarks

When this flag is True, the program automatically updates bridge objects before running an analysis if it detects anything has been changed that might affect the bridge analysis.

This flag is by default set to True for each new Sap2000 Object.

## VBA Example



```
Sub GetBridgeAnalysisFlag()

       'dimension variables

            Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim Flag as Boolean



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



       'get flag

          Flag = SapModel.BridgeObj.GetBridgeUpdateForAnalysisFlag



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.00.

This function is DEPRECATED as of version v26.0.0. Please use cBridgeModeler\_1 which is documented in the CSI\_OAPI\_Bridge\_Modeler.chm help file.This topic is maintained for reference.

## See Also

[SetBridgeUpdateForAnalysisFlag](SetBridgeUpdateForAnalysisFlag.md)
