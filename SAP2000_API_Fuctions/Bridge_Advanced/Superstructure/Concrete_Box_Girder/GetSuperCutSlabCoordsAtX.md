# GetSuperCutSlabCoordsAtX

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutSlabCoordsAtX

## VB6 Procedure

Function GetSuperCutSlabCoordsAtX(ByVal Name As String, ByVal CutIndex As Long, ByVal X As Double, ByRef Status As Long, ByRef y1 As Double, ByRef y2 As Double, ByRef y3 As Double, ByRef y4 As Double) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of the section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by function CountSuperCut. Section cuts will be in order of increasing Station and increasing SuperCutType.

X

The X coordinate in the section local coordinate system at which a vertical line is passed through the section and the slab coordinates are returned.

Status

This is 0, 1, or 2.

0 = No portion of the section is cut.

1 = Only the section is cut; no interior cell is cut.

2 = The section and an interior cell are cut.

y1

The topmost Y coordinate where the vertical line cuts the section.  This item is returned as zero when Status < 1.

y2

The bottommost Y coordinate where the vertical line cuts the section.  This item is returned as zero when Status < 1.

y3

The topmost Y coordinate where the vertical line cuts an interior cell.  This item is returned as zero when Status < 2.

y4

The bottommost Y coordinate where the vertical line cuts an interior cell.  This item is returned as zero when Status < 2.

## Remarks

This function returns information about the box girder slab thicknesses at a given horizontal location across the box girder section.

The function returns zero if the information is successfully retrieved; otherwise it returns a nonzero value.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutSlabCoordsAtX()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Status As Long

      Dim y1 As Double

      Dim y2 As Double

      Dim y3 As Double

      Dim y4 As Double



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



   'get slab thickness information at X = 60

      ret = SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutSlabCoordsAtX("BOBJ1", 1, 60, Status, y1, y2, y3, y4)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also
