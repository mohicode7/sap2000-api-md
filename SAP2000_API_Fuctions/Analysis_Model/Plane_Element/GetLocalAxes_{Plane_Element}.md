# GetLocalAxes

## Syntax

SapObject.SapModel.PlaneElm.GetLocalAxes

## VB6 Procedure

Function GetLocalAxes(ByVal Name As String, ByRef Ang As Double) As Long

## Parameters

Name

The name of an existing plane element.

Ang

This is the angle that the local 1 and 2 axes are rotated about the positive local 3 axis from the default orientation. The rotation for a positive angle appears counter clockwise when the local +3 axis is pointing toward you. [deg]

## Remarks

This function retrieves the local axis angle assignment for plane elements.

The function returns zero if the assignment is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPlaneElementLocalAxisAngle()

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



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 3-001-incomp.sdb")



   'assign area object local axis angle

      ret = SapModel.AreaObj.SetLocalAxes("3", 30)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get plane element local axis angle

      ret = SapModel.PlaneElm.GetLocalAxes("3", Ang)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
