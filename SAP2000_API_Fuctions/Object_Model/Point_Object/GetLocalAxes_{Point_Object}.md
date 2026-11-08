# GetLocalAxes

## Syntax

SapObject.SapModel.PointObj.GetLocalAxes

## VB6 Procedure

Function GetLocalAxes(ByVal Name As String, ByRef a As Double, ByRef b As Double, ByRef c As Double, ByRef Advanced As Boolean) As Long

## Parameters

Name

The name of an existing point object.

a, b, c

The local axes of the point are defined by first setting the positive local 1, 2 and 3 axes the same as the positive global X, Y and Z axes and then doing the following: [deg]

1.    Rotate about the 3 axis by angle a.

2.    Rotate about the resulting 2 axis by angle b.

3.    Rotate about the resulting 1 axis by angle c.

Advanced

This item is True if the point object local axes orientation was obtained using advanced local axes parameters.

## Remarks

This function retrieves the local axes angles for a point object.

The function returns zero if the local axes angles are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointLocalAxes()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim a As Double, b As Double, c As Double

      Dim Advanced As Boolean



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



   'get local axes assignments

      ret = SapModel.PointObj.GetLocalAxes("1", a, b, c, Advanced)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetLocalAxes](SetLocalAxes_%7BPoint_Object%7D.md)
