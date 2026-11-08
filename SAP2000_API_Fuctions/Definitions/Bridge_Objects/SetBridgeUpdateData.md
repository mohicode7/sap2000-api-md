# SetBridgeUpdateData (Note:  Newer function available)

## Syntax

SapObject.SapModel.BridgeObj.SetBridgeUpdateData

## VB6 Procedure

Function SetBridgeUpdateData(ByVal Name As String, ByVal Action As Long, ByVal ModelType As Long, ByVal MaxDeckSegLength As Double, ByVal MaxCapSegLength As Double, ByVal MaxColSegLength As Double, ByVal SubMeshSize As Double) As Long

## Parameters

Name

The name of an existing bridge object.

Action

This is 1, 2 or 3, indicating the action to be taken.

1 = Update linked model.

2 = Clear all from linked model.

3 = Convert to unlinked model.

ModelType

This is 1, 2 or 3, indicating the linked bridge model type. This item applies only when the Action item is 1 (update linked model).

1 = Spine model (frame).

2 = Area model.

3 = Solid model.

MaxDeckSegLength

The maximum length for the deck objects in the linked bridge model. This item applies only when the Action item is 1 (update linked model). [L]

MaxCapSegLength

The maximum length for the cap beam objects in the linked bridge model. This item applies only when the Action item is 1 (update linked model). [L]

MaxColSegLength

The maximum length for the column objects in the linked bridge model. This item applies only when the Action item is 1 (update linked model). [L]

SubMeshSize

The maximum submesh size for area and solid objects in the linked bridge model. This item applies only when the Action item is 1 (update linked model) and the ModelType item is 2 or 3 (area or solid model). [L]

## Remarks

This function updates a linked bridge model, clears all objects from a linked bridge model, or converts a linked bridge model to an unlinked model.

The function returns zero if it is successful.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a bridge object named BOBJ1 in it.



```
Sub SetBridgeObjectUpdateData()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

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



       'update linked bridge model

          ret = SapModel.BridgeObj.SetBridgeUpdateData("BOBJ1", 1, 2, 150, 150, 150, 50)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 14.20.

This function is DEPRECATED as of version v26.0.0. Please use cBridgeModeler\_1 which is documented in the CSI\_OAPI\_Bridge\_Modeler.chm help file.This topic is maintained for reference.

## See Also

[GetBridgeUpdateData](GetBridgeUpdateData.md)
