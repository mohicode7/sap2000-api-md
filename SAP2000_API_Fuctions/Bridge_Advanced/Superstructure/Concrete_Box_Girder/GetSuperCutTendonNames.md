# GetSuperCutTendonNames

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutTendonNames

## VB6 Procedure

Function GetSuperCutTendonNames(ByVal Name As String, ByVal CutIndex As Long, ByVal TendonIndex As Long, ByRef BridgeTendon As String, ByRef TendonObj as String) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of the section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by the function CountSuperCut. Section cuts will be in order of increasing Station and increasing SuperCutType.

TendonIndex

The index number of a tendon in this section cut of this bridge object. This must be from 0 to CountTendon-1, where CountTendon is the number of tendons returned by function GetSuperCutSectionValues using Item = 12.

BridgeTendon

The name of the tendon inside of the bridge object corresponding to TendonIndex.

TendonObj

The name of the tendon object created by the program from the bridge object tendon corresponding to TendonIndex.

## Remarks

This function returns the name of a single tendon object, giving access to tendon assignments, tendon section, and material property.

The function returns zero if the information is successfully retrieved; otherwise it returns a nonzero value.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutTendonNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

      Dim CountTendon As Long

      Dim BridgeTendon As String

      Dim TendonObj As String



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



   'get tendon count at section cut 1

      ret = SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutSectionValues ("BOBJ1", 12, 1, CountTendon)



   'get tendon object name for first tendon (0)

      ret = SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutTendonNames("BOBJ1", 1, 0, BridgeTendon, TendonObj)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[CountSuperCut](../CountSuperCut.md)

[GetSuperCutSectionValues](GetSuperCutSectionValues.md)

[GetSuperCutTendonValues](GetSuperCutTendonValues.md)
