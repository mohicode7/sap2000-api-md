# GetSuperCutWebCoordsAtY

## Syntax

SapObject.SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutWebCoordsAtY

## VB6 Procedure

Function GetSuperCutWebCoordsAtY(ByVal Name As String, ByVal CutIndex As Long, ByVal Y As Double, ByRef NumberWebs As Long, ByRef WebIsCut() As Boolean, ByRef WebLeft() As Double, ByRef WebRight() As Double) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of the section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by the function CountSuperCut. Section cuts will be in order of increasing Station and increasing SuperCutType.

Y

The Y coordinate in the section local coordinate system at which the web coordinates are returned.

NumberWebs

The number of webs in the section.

WebIsCut

This is a array of booleans indicating if each web is cut by a horizontal line at the specified Y coordinate.

WebLeft

This is a array of X coordinates of the left side of each web. If the web is not cut by a horizontal line at the specified Y coordinate, this value is returned as zero.

WebRight

This is a array of X coordinates of the right side of each web. If the web is not cut by a horizontal line at the specified Y coordinate, this value is returned as zero.

## Remarks

This function returns information about the box girder web thicknesses at a given elevation in the box girder section.

The function returns zero if the information is successfully retrieved; otherwise it returns a nonzero value.

If the bridge object is not currently linked to existing objects in the model, an error is returned.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it.



```
Sub GetBridgeSuperCutWebCoordsAtY()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberWebs As Long

      Dim WebIsCut() As Boolean

      Dim WebLeft() As Double

      Dim WebRight() As Double



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



   'get web information

      ret = SapModel.BridgeAdvancedSuper.BASConcBox.GetSuperCutWebCoordsAtY("BOBJ1", 1, 44, NumberWebs, WebIsCut, WebLeft, WebRight)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also
