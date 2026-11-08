# GetSuperCutWebValues

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutWebValues

## VB6 Procedure

Function GetSuperCutWebValues(ByVal Name As String, ByVal CutIndex As Long, ByVal WebIndex As Long, ByVal Item as Long, ByRef Value as Double) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of the section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by the function CountSuperCut. Section cuts will be in order of increasing Station and increasing SuperCutType.

WebIndex

The index number of a web in this section cut of this bridge object. This must be from 0 to CountWeb-1, where CountWeb is the number of webs returned by thenfunction GetSuperCutSectionValues using Item = 1. Webs count from left to right when looking upstation.

Item

This is an integer from 1 to 6, inclusive, indicating the type of property value to be gotten:

1 = Angle from vertical (Y) axis, clockwise is positive

2 = Minimum horizontal (X) web thickness

3 = Minimum top slab thickness above cell to left of web

4 = Minimum bottom slab thickness above cell to left of web

5 = Top width of cell to left of web measured from centerline of girders on each side of cell

6 = Bottom width of cell to left of web measured from centerline of girders on each side of cell.

Value

The value of the requested item:

1 = Angle from vertical (Y) axis, clockwise is positive

Abs(Value) < 90. [deg]

2 = Minimum horizontal (X) web thickness

Value > 0. [L]

3 = Minimum top slab thickness

Value < 0. [L]

4 = Minimum bottom slab thickness

Value > 0. [L]

5 = Top width of cell

Value >= 0. [L]

6 = Bottom width of cell

Value >= 0. [L]

## Remarks

This function returns an individual section property for a single web at a single superstructure section cut in a bridge object. These properties are calculated for the section before skew, grade, and superelevation are applied. Coordinate values are measured from the lower-left corner of the section bounding-box. X is positive to the right when looking upstation, and Y is positive upward

The function returns zero if the information is successfully retrieved; otherwise it returns a nonzero value. An error is returned for items 3, 4, 5 and 6 if the WebIndex is specified as 0.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutWebValues()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim Count As Long

          Dim CountWeb As Long

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



       'get web count at section cut 1

          ret = SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutSectionValues ("BOBJ1", 1, 1, CountWeb)



       'get minimum thickness for first web (0)

          ret = SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutWebValues("BOBJ1", 1, 0, 2, Value)



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
