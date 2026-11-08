# GetLocalAxes

## Syntax

SapObject.SapModel.AreaElm.GetLocalAxes

## VB6 Procedure

Function GetLocalAxes(ByVal Name As String, ByRef Ang As Double) As Long

## Parameters

Name

The name of an existing area element.

Ang

This is the angle that the local 1 and 2 axes are rotated about the positive local 3 axis from the default orientation. The rotation for a positive angle appears counter clockwise when the local +3 axis is pointing toward you. [deg]

## Remarks

This function retrieves the local axis angle assignment for area elements.

The function returns zero if the assignment is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaElementLocalAxisAngle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Ang As Double



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



   'assign area object local axis angle

      ret = SapModel.AreaObj.SetLocalAxes("3", 30)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get area element local axis angle

      ret = SapModel.AreaElm.GetLocalAxes("3", Ang)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
