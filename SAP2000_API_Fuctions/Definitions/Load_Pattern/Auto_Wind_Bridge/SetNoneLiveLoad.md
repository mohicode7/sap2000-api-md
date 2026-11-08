# SetNoneLiveLoad

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.SetNoneLiveLoad

## VB6 Procedure

Function SetNoneLiveLoad(ByVal Name As String) As Long

## Parameters

Name

The name of an existing bridge wind - live load type pattern.

## Remarks

This function sets the auto wind – live load type for the specified load pattern to None.

The function returns zero if the loading type is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignWindLiveLoadNone()

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



   'set auto wind live load loading type to None

      ret = SapModel.LoadPatterns.AutoWindBridge.SetNoneLiveLoad("WINDLIVE")
   

   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v21.0.0.
