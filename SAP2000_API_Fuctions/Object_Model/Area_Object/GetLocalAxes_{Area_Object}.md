# GetLocalAxes

## Syntax

SapObject.SapModel.AreaObj.GetLocalAxes

## VB6 Procedure

Function GetLocalAxes(ByVal Name As String, ByRef Ang As Double, ByRef Advanced As Boolean) As Long

## Parameters

Name

The name of an existing area object.

Ang

This is the angle that the local 1 and 2 axes are rotated about the positive local 3 axis from the default orientation. The rotation for a positive angle appears counter clockwise when the local +3 axis is pointing toward you. [deg]

Advanced

This item is True if the area object local axes orientation was obtained using advanced local axes parameters.

## Remarks

This function retrieves the local axis angle assignment for area objects.

The function returns zero if the assignment is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectLocalAxisAngle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Ang As Double

      Dim Advanced As boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'get area object local axis angle

      ret = SapModel.AreaObj.GetLocalAxes("3", Ang, Advanced)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetLocalAxes](SetLocalAxes_%7BArea_Object%7D.md)
