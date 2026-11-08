# GetSuperCutSectionPropsAtY

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutSectionPropsAtY

## VB6 Procedure

Function GetSuperCutSectionPropsAtY (ByVal Name As String, ByVal CutIndex As Long, ByVal Y as Double, ByVal AboveY as Boolean, ByRef Ycg as Double, ByRef Area as Double, ByRef Inertia as Double) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of the section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by the function CountSuperCut. Section cuts will be in order of increasing Station and increasing SuperCutType.

Y

The Y coordinate in the section local coordinate system above or below which the section properties are calculated.

AboveY

Boolean indicating whether properties are to be computed for the region above (if true) or below (if false) the specified coordinate Y.

Ycg

Y coordinate of the centroid of the region above/below specified coordinate Y. [L]

Area

Area of the region above/below specified coordinate Y. [L2]

Inertia

Moment of inertia of the region above/below specified coordinate Y, taken about a horizontal axis at Ycg. [L4]

## Remarks

This function returns section properties for the region above or below a given Y coordinate value at a single superstructure section cut in a bridge object. These properties are calculated for the section before skew, grade, and superelevation are applied. Coordinate values are measured from the lower-left corner of the section bounding-box. X is positive to the right when looking upstation, and Y is positive upward.

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

      Dim Y As Double

      Dim Ycg As Double

      Dim Area As Double

      Dim Inertia As Double



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



   'get section properties for region above Y = 10 at first cut (0)

      ret = SapModel.BridgeAdvancedSuper.BASConcBox. GetSuperCutSectionPropsAtY ("BOBJ1", 0, Y, True, Ycg, Area, Inertia)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.0.0.

## See Also

[CountSuperCut](../CountSuperCut.md)
