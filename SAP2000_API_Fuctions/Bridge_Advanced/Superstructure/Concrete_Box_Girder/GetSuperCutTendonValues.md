# GetSuperCutTendonValues

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutTendonValues

## VB6 Procedure

Function GetSuperCutTendonValues(ByVal Name As String, ByVal CutIndex As Long, ByVal TendonIndex As Long, ByVal Item as Long, ByRef Value as Double) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of the section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by the function CountSuperCut. Section cuts will be in order of increasing Station and increasing SuperCutType.

TendonIndex

The index number of a tendon in this section cut of this bridge object. This must be from 0 to CountTendon-1, where CountTendon is the number of tendons returned by the function GetSuperCutSectionValues using Item = 12.

Item

This is an integer from 1 to 4, inclusive, indicating the type of property value to be gotten:

1 = X coordinate of tendon centroid, Xcg

2 = Y coordinate of tendon centroid, Ycg

3 = Duct diameter for tendon

4 = Bonding type for tendon

5 = Tendon slope

Value

The value of the requested item:

1 = X coordinate of tendon centroid, Xcg

Any value OK. [L]

2 = Y coordinate of tendon centroid, Ycg

Any value OK. [L]

3 = Duct diameter for tendon

Value >= 0. [L]

4 = Bonding type for tendon

1 = Bonded

2 = Unbonded

5 = Tendon slope

Any value OK. [L/L]

## Remarks

This function returns an individual section property for a single tendon at a single superstructure section cut in a bridge object. These properties are calculated for the section before skew, grade, and superelevation are applied. Coordinate values are measured from the lower-left corner of the section bounding-box. X is positive to the right when looking upstation, and Y is positive upward.

The function returns zero if the information is successfully retrieved; otherwise it returns a nonzero value.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutTendonValues()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

      Dim CountTendon As Long

      Dim Value As Double



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



   'get Y coordinate of centroid for first tendon (0)

      ret = SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutTendonValues("BOBJ1", 1, 0, 2, Value)



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

[GetSuperCutTendonNames](GetSuperCutTendonNames.md)
