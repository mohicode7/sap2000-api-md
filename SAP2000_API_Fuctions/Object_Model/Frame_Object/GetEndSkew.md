# GetEndSkew

## Syntax

SapObject.SapModel.FrameObj.GetEndSkew

## VB6 Procedure

Function GetEndSkew(ByVal Name As String, ByRef SkewI As Double, ByRef SkewJ As Double) As Long

## Parameters

Name

The name of an existing frame object.

SkewI

The angle in degrees measured counter clockwise from the positive local 3-axis to a line parallel to the I-End of the frame object (-90 < SkewI < 90). [deg]

SkewJ

The angle in degrees measured counter clockwise from the positive local 3-axis to a line parallel to the J-End of the frame object (-90 < SkewJ < 90). [deg]

## Remarks

This function retrieves frame object end skew assignments.

The function returns zero if the end skew data is successfully retrieved, otherwise it returns a nonzero value.

End skew assignments are only applicable to straight frame objects. An error is returned if skew data is requested for a curved frame object.

End skew data is only used in the program to plot the extruded view of bridge objects that have been updated as spine models.

## VBA Example



```
Sub GetFrameEndSkewData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SkewI As Double

      Dim SkewJ As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'get end skew data

      ret = SapModel.FrameObj.GetEndSkew("15", SkewI, SkewJ)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetEndSkew](SetEndSkew.md)
