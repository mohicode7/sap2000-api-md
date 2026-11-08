# CountSuperCut

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.CountSuperCut

## VB6 Procedure

Function CountSuperCut(ByVal Name As String, ByRef Count As Long) As Long

## Parameters

Name

The name of an existing bridge object.

Count

The number of section cuts in this bridge object. They will be identified in subsequent API functions using the indices 0 to Count-1. There may be one or two section cuts at each output station along the length of the superstructure.

## Remarks

This function returns the number of superstructure section cuts that are available for getting analysis results and performing design.

The function returns zero if the count is successfully retrieved, otherwise it returns a nonzero value.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutCount()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

      Dim FileName As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open an existing file

      FileName = "C:\SapAPI\MyBridge.sdb"

      ret = SapModel.File.OpenFile(FileName)



   'get section cut count

      ret = SapModel.BridgeAdvancedSuper.CountSuperCut("BOBJ1", Count)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also
