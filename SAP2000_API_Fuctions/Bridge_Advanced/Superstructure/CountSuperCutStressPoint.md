# CountSuperCutStressPoint

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.CountSuperCutStressPoint

## VB6 Procedure

Function CountSuperCutStressPoint(ByVal Name As String, ByVal CutIndex As Long, ByRef Count As Long) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by function CountSuperCut.

Count

The number of stress points for this section cut in this bridge object. They will be identified in subsequent API functions using the indices 0 to Count-1.

## Remarks

This function returns the number of stress points at the specified superstructure section cut.

The function returns zero if the count is successfully retrieved, otherwise it returns a nonzero value.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutStressPointCount()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

      Dim CountPoint As Long



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



   'get section cut stress point count

      ret = SapModel.BridgeAdvancedSuper.CountSuperCutStressPoint("BOBJ1", 1, CountPoint)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[CountSuperCut](CountSuperCut.md)
