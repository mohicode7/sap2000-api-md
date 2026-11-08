# LinesParallelToCoordAxis

## Syntax

SapObject.SapModel.SelectObj.LinesParallelToCoordAxis

## VB6 Procedure

Function LinesParallelToCoordAxis(ByRef ParallelTo() As Boolean, Optional ByVal CSys As String = "Global", Optional ByVal Tolerance As Double = 0.057, Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

ParallelTo

This is an array of six booleans representing three coordinate axes and three coordinate planes. Any combination of the six may be specified.

ParallelTo(0) = X axis

ParallelTo(1) = Y axis

ParallelTo(2) = Z axis

ParallelTo(3) = XY plane

ParallelTo(4) = XZ plane

ParallelTo(5) = YZ plane

CSys

The name of the coordinate system to which the ParallelTo items apply.

Tolerance

Line objects that are within this angle in degrees of being parallel to a specified coordinate axis or plane are selected or deselected. [deg]

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects objects parallel to specified coordinate axes or planes.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectParallelToAxes()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim ParallelTo() As Boolean



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



   'select parallel to axes

      ReDim ParallelTo(5)

      ParallelTo(2) = True

      ret = SapModel.SelectObj.LinesParallelToCoordAxis(ParallelTo)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[LinesParallelToLine](LinesParallelToLine.md)
