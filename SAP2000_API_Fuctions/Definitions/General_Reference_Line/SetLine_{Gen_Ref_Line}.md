# SetLine

## Syntax

SapObject.SapModel.GenRefLine.SetLine

## VB6 Procedure

Function SetLine(ByVal Name As String, ByVal DiscLength As Double, ByVal DiscAngle As Double, Optional ByVal CSys As String = "Global", Optional ByVal Color As Long = -1, Optional ByVal Visible As Boolean = True) As Long

## Parameters

Name

This is the name of a general reference line. If this is the name of an existing general reference line, that general reference line is modified; otherwise a new general reference line is added.

DiscLength

The maximum segment discretization length of the segments used to define curves in the general reference line. [L]

DiscAngle

The maximum discretization angle in degrees for the general reference line. [deg]

CSys

The name of the coordinate system in which the general reference line is defined.

Color

The display color assigned to the general reference line. If Color is specified as -1, the program will automatically assign a color.

Visible

The item is True if the general reference line should be displayed in windows displaying the model.

## Remarks

The function returns zero if the general reference line is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddGenRefLine()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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


   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.0.

## See Also

[GetLine](GetLine_%7BGen_Ref_Line%7D.md) [SetLinePlanPoints](SetLinePlanPoints.md) [SetLineElevPoints](SetLineElevPoints.md)

[Delete](Delete_%7BGen_Ref_Line%7D.md)
