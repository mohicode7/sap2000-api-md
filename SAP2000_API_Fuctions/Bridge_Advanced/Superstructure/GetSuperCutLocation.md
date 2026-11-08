# GetSuperCutLocation

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.GetSuperCutLocation

## VB6 Procedure

Function GetSuperCutLocation(ByVal Name As String, ByVal CutIndex As Long, ByRef Location As Long, ByRef Station As Double, ByRef XRefPt As Double, ByRef YRefPt As Double, ByRef Skew As Double, ByRef Grade As Double, ByRef SuperElev As Double) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by function CountSuperCut. Section cuts will be in order of increasing Station and increasing SuperCutType.

Location

This is 1 or 2, indicating whether the CutIndex section cut occurs before or after the associated station.

1 = Before the specified station.

2 = After the specified station.

Station

The station ordinate of the CutIndex section cut at the reference line of the superstructure. [L]

XRefPt, YRefPt

The transverse and vertical coordinates in the section of the reference point that corresponds to the layout line in the bridge object. XRefPt is positive to the right when looking upstation. YRefPt is positive upward. Coordinates are measured from the lower-left corner of the section bounding-box before skew, grade, and superelevation are applied. The rotations of the section due to skew, grade, and superelevation occur about the reference point. [L]

Skew

The skew angle, in degrees, of the section cut, measured from the horizontal normal to the superstructure reference line, with positive being about the upward vertical axis.

Grade

The grade, as a slope (abs(Grade) < 1.0), giving the vertical rise per distance along the superstructure reference line.

SuperElev

The superelevation, as a slope (abs(SuperElev) < 1.0), giving the vertical rise per distance along the transverse normal to the superstructure reference line.

## Remarks

This function returns location and orientation information about a single superstructure section cut in a bridge object.

The function returns zero if the information is successfully retrieved, otherwise it returns a nonzero value.

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

          Dim Location As Long

          Dim Station As Double

          Dim XRefPt As Double,

          Dim YRefPt As Double

          Dim Skew As Double

          Dim Grade As Double

          Dim SuperElev As Double



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



       'get section cut location information

          ret = SapModel.BridgeAdvancedSuper.GetSuperCutLocation("BOBJ1", 1, Location, Station, RefPt, YRefPt Skew, Grade, SuperElev)



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
