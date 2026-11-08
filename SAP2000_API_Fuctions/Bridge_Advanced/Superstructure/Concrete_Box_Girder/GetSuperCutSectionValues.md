# GetSuperCutSectionValues

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutSectionValues

## VB6 Procedure

Function GetSuperCutSectionValues (ByVal Name As String, ByVal CutIndex As Long, ByVal Item as Long, ByRef Value as Double) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of the section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by the function CountSuperCut. Section cuts will be in order of increasing Station and increasing SuperCutType.

Item

This is an integer from 1 to 12, inclusive, indicating the type of property value to be gotten:

1 = Number of girders or webs

2 = Design area of top slab, ASlabTop

3 = Design area of bottom slab, ASlabBot

4 = Width of top slab, BSlabTop

5 = Width of bottom slab, BSlabBot

6 = X coordinate of top slab centroid, XSlabTop

7 = X coordinate of bottom slab centroid, XSlabBot

8 = Y coordinate of top slab centroid, YSlabTop

9 = Y coordinate of bottom slab centroid, YSlabBot

10 = Area inside torsion circuit, ATorsion

11 = Length of torsion circuit, LTorsion

12 = Number of tendons

13 = Top outside width of torsion circuit, BTorsionTop

14 = Bottom outside width of torsion circuit, BTorsionBot

15 = Left outside length of torsion circuit (along slope), HTorsionLeft

16 = Right outside length of torsion circuit (along slope), HTorsionRight

Value

The value of the requested item:

1 = Number of girders or webs

Value >= 0, integral.

2 = Design area of top slab, ASlabTop

Value > 0. [L2]

3 = Design area of bottom slab, ASlabBot

Value >= 0. [L2]

4 = Width of top slab, BSlabTop

Value > 0. [L]

5 = Width of bottom slab, BSlabBot

Value >= 0. [L]

6 = X coordinate of top slab centroid, XSlabTop

Any value is valid. [L]

7 = X coordinate of bottom slab centroid, XSlabBot

Any value is valid. [L]

8 = Y coordinate of top slab centroid, YSlabTop

Value >= 0. [L]

9 = Y coordinate of bottom slab centroid, YSlabBot

Value <= 0. [L]

10 = Area inside torsion circuit, ATorsion

Value >= 0. [L2]

11 = Length of torsion circuit, LTorsion

Value >= 0. [L]

12 = Number of tendons

Value >= 0.

13 = Top outside width of torsion circuit, BTorsionTop

Value > 0. [L]

14 = Bottom outside width of torsion circuit, BTorsionBot

Value > 0. [L]

15 = Left outside length of torsion circuit (along slope), HTorsionLeft

Value > 0. [L]

16 = Right outside length of torsion circuit (along slope), HTorsionRight

Value > 0. [L]

## Remarks

This function returns an individual section property at a single superstructure section cut in a bridge object. These properties are calculated for the section before skew, grade, and superelevation are applied. Coordinate values are measured from the lower-left corner of the section bounding-box. X is positive to the right when looking upstation, and Y is positive upward.

The function returns zero if the information is successfully retrieved; otherwise it returns a nonzero value.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutSectionValues()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

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



   'get section cut section top-slab area in Value

      ret = SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutSectionValues ("BOBJ1", 1, 2, Value)



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
