# GetBridgeUpdateData (Note:  Newer function available)

## Syntax

SapObject.SapModel.BridgeObj.GetBridgeUpdateData

## VB6 Procedure

Function GetBridgeUpdateData(ByVal Name As String, ByRef LinkedModelExists As Boolean, ByRef ModelType As Long, ByRef MaxDeckSegLength As Double, ByRef MaxCapSegLength As Double, ByRef MaxColSegLength As Double, ByRef SubMeshSize As Double) As Long

## Parameters

Name

The name of an existing bridge object.

LinkedModelExists

This item is True if a linked bridge model exists for the specified bridge object.

ModelType

This is 1, 2 or 3, indicating the linked bridge model type. This item applies only when the LinkedModelExists item is True.

1 = Spine model (frame).

2 = Area model.

3 = Solid model.

MaxDeckSegLength

The maximum length for the deck objects in the linked bridge model. This item applies only when the LinkedModelExists item is True. [L]

MaxCapSegLength

The maximum length for the cap beam objects in the linked bridge model. This item applies only when the LinkedModelExists item is True. [L]

MaxColSegLength

The maximum length for the column objects in the linked bridge model. This item applies only when the LinkedModelExists item is True. [L]

SubMeshSize

The maximum submesh size for area and solid objects in the linked bridge model. This item applies only when the LinkedModelExists item is True and the ModelType item is 2 or 3 (area or solid model). [L]

## Remarks

This function returns a flag indicating if the specified bridge object is currently linked to existing objects in the model. If the bridge object is linked, it returns the model type (spine, area or solid) and meshing data used when the linked bridge model was updated.

The function returns zero if the information is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a bridge object named BOBJ1 in it.



```
Sub GetBridgeObjectUpdateData()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim FileName As String

          Dim LinkedModelExists As Boolean

          Dim ModelType As Long

          Dim MaxDeckSegLength As Double

          Dim MaxCapSegLength As Double

          Dim MaxColSegLength As Double

          Dim SubMeshSize As Double



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



       'get bridge update data

          ret = SapModel.BridgeObj.GetBridgeUpdateData("BOBJ1", LinkedModelExists, ModelType, MaxDeckSegLength, MaxCapSegLength, MaxColSegLength, SubMeshSize)



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

[SetBridgeUpdateData](SetBridgeUpdateData.md)
