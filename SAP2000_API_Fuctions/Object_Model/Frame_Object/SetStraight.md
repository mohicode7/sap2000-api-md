# SetStraight

## Syntax

SapObject.SapModel.FrameObj.SetStraight

## VB6 Procedure

Function SetStraight(ByVal Name As String) As Long

## Parameters

Name

The name of a defined curved frame object.

## Remarks

This function sets a curved frame object straight.

The function returns zero if the frame object type is successfully set, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameStraight()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'set frame curved

      ret = SapModel.FrameObj.SetCurved("15", 2, -200, 0, 300, "", 0, 16)



   'set frame straight

      ret = SapModel.FrameObj.SetStraight("15")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetCurved](SetCurved.md)

[GetCurved](../../Obsolete_Functions/GetCurved.md)

[GetType](GetType_%7BFrame_Object%7D.md)
