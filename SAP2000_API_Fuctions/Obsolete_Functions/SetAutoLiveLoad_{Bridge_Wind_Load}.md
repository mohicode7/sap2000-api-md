# SetAutoLiveLoad {Bridge Wind Load} (Note:  Newer function available)

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.SetAutoLiveLoad

## VB6 Procedure

Function SetAutoLiveLoad(ByVal Name As String, ByVal RefLoadPat As String, ByVal Height As Double) As Long

## Parameters

Name

The name of an existing bridge wind - live load type pattern.

RefLoadPat

The name of an existing bridge wind load pattern that is referenced from this wind on live load pattern.

Height

The height above the roadway surface at which the wind on live load should be applied. [L]

## Remarks

This function applies auto wind on live load parameters.

The function returns zero if the parameters are successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAutoWindLiveLoad()

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
  
 
   'open existing model containing a bridge object

          ret = SapModel.File.OpenFile(“C:\Temp\BridgeModel.bdb” )



       'add new load pattern

          ret = SapModel.LoadPatterns.Add("WIND", LTYPE_WINDONLIVELOAD)



       'set auto wind live load loading type to Auto

          ret = SapModel.LoadPatterns.AutoWindBridge.SetAutoLiveLoad("WINDLIVE", "WIND", 0)
  
   

       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
```



## Release Notes

Initial release in v21.0.0.

This function is obsolete and was superseded by [SetAutoLiveLoad\_1](../Definitions/Load_Pattern/Auto_Wind_Bridge/SetAutoLiveLoad_1%7BBridge_Wind_Load%7D.md) in release v25.2.0. This function is maintained for backwards compatibility.

## See Also

[GetAutoLiveLoad](GetAutoLiveLoad_%7BBridge_Wind_Load%7D.md)

[SetAutoLiveLoad\_1](../Definitions/Load_Pattern/Auto_Wind_Bridge/SetAutoLiveLoad_1%7BBridge_Wind_Load%7D.md)
