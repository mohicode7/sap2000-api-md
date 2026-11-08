# GetLocalAxes

## Syntax

SapObject.SapModel.PointElm.GetLocalAxes

## VB6 Procedure

Function GetLocalAxes(ByVal Name As String, ByRef a As Double, ByRef b As Double, ByRef c As Double) As Long

## Parameters

Name

The name of an existing point element.

a, b, c

The local axes of the point are defined by first setting the positive local 1, 2 and 3 axes the same as the positive global X, Y and Z axes and then doing the following: [deg]

1.    Rotate about the 3 axis by angle a.

2.    Rotate about the resulting 2 axis by angle b.

3.    Rotate about the resulting 1 axis by angle c.

## Remarks

This function retrieves the local axes angles for a point element.

The function returns zero if the local axes angles are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointElmLocalAxes()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim a As Double, b As Double, c As Double



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



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get local axes assignments

      ret = SapModel.PointElm.GetLocalAxes("1", a, b, c)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
