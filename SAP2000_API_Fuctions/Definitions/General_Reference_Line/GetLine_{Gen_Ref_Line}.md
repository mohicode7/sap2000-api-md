# GetLine

## Syntax

SapObject.SapModel.GenRefLine.GetLine

## VB6 Procedure

Function GetLine(ByVal Name As String, ByRef DiscLength As Double, ByRef DiscAngle As Double, ByRef Color As Long, ByRef Visible As Boolean) As Long

## Parameters

Name

The name of an existing general reference line.

DiscLength

The maximum segment discretization length of the segments used to define curves in the general reference line. [L]

DiscAngle

The maximum discretization angle in degrees for the general reference line. [deg]

Color

The display color assigned to the general reference line.

Visible

Specifies whether the general reference line will be displayed in windows displaying the model.

## Remarks

The function returns zero if the general reference line data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetGenRefLine()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim discLength As Double, discAngle As Double

      Dim color As Long

      Dim visible As Boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'define new general reference line

      ret = SapModel.GenRefLine.SetLine("GRef1", 120, 5)



   'get new general reference line

      ret = SapModel.GenRefLine.GetLine("GRef1", discLength, discAngle, color, visible)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.0.

## See Also

[SetLine](SetLine_%7BGen_Ref_Line%7D.md) [GetLinePlanPoints](GetLinePlanPoints.md) [GetLineElevPoints](GetLineElevPoints.md)
