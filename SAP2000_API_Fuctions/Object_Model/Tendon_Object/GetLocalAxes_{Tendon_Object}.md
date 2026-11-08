# GetLocalAxes

## Syntax

SapObject.SapModel.TendonObj.GetLocalAxes

## VB6 Procedure

Function GetLocalAxes(ByVal Name As String, ByRef Ang As Double) As Long

## Parameters

Name

The name of an existing tendon object.

Ang

This is the angle that the local 2 and 3 axes are rotated about the positive local 1 axis, from the default orientation. The rotation for a positive angle appears counter clockwise when the local +1 axis is pointing toward you. [deg]

## Remarks

This function retrieves the tendon local axis angle assignment for tendon objects.

The function returns zero if the assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonLocalAxisAngle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Ang As Double

      Dim Name As String



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



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'assign tendon local axis angle

      ret = SapModel.TendonObj.SetLocalAxes(Name, 30)



   'get tendon local axis angle

      ret = SapModel.TendonObj.GetLocalAxes(Name, Ang)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetLocalAxes](SetLocalAxes_%7BTendon_Object%7D.md)
