# GetAutoLiveLoad\_1 {Bridge Wind Load}

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.GetAutoLiveLoad\_1

## VB6 Procedure

Function GetAutoLiveLoad\_1(ByVal Name As String, ByRef RefLoadPat As String, ByRef Height As Double, ByRef LLHeight As Double) As Long

## Parameters

Name

The name of an existing bridge wind - live load type pattern.

RefLoadPat

The name of an existing bridge wind load pattern that is referenced from this wind on live load pattern.

Height

The height above the roadway surface at which the wind on live load should be applied. [L]

LLHeight

The height of the live load surface to which the wind should be applied. This parameter does not apply for AASHTO 2018 or AASHTO 2020. [L]

## Remarks

This function retrieves auto wind on live load parameters.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAutoWindLiveLoad_1()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long
 
      Dim RefLoadPat As String
      Dim Height As Double
      Dim LLHeight As Double
 
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

          ret = SapModel.LoadPatterns.AutoWindBridge.SetAutoLiveLoad_1("WINDLIVE", "WIND", 0, 2.5)
  
 
   'get auto wind live load parameters
      ret = SapModel.LoadPatterns.AutoWindBridge.GetAutoLiveLoad_1("WINDLIVE", RefLoadPat, Height, LLHeight)
   

       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
```



## Release Notes

Initial release in v25.2.0.

This function supersedes [GetAutoLiveLoad](../../../Obsolete_Functions/GetAutoLiveLoad_%7BBridge_Wind_Load%7D.md).

## See Also

[SetAutoLiveLoad\_1](SetAutoLiveLoad_1%7BBridge_Wind_Load%7D.md)
