# Align

## Syntax

SapObject.SapModel.EditPoint.Align

## VB6 Procedure

Function Align(ByVal MyType As Long, ByVal Ordinate As Double, ByRef NumberPoints As Long, ByRef PointName() As String) As Long

## Parameters

MyType

This is 1, 2, 3 or 4, indicating the alignment option.

1 = Align points to X-ordinate in present coordinate system

2 = Align points to Y-ordinate in present coordinate system

3 = Align points to Z-ordinate in present coordinate system

4 = Align points to nearest selected line object, area object edge or solid object edge

Ordinate

The X, Y or Z ordinate that applies if MyType is 1, 2 or 3, respectively. [L]

NumberPoints

The number of point objects that are in a new location after the alignment is complete.

PointName

This is an array of the name of each point object that is in a new location after the alignment is complete.

## Remarks

This function aligns selected point objects.

The function returns zero if the alignment is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub AlignPointObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberPoints As Long

      Dim PointName() As String



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



   'align point objects

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.PointObj.SetSelected("1", True)

      ret = SapModel.PointObj.SetSelected("2", True)

      ret = SapModel.PointObj.SetSelected("3", True)

      ret = SapModel.EditPoint.Align(1, -300, NumberPoints, PointName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
