# CoordinateRange

## Syntax

SapObject.SapModel.SelectObj.CoordinateRange

## VB6 Procedure

Function CoordinateRange(ByVal XMin As Double, ByVal XMax As Double, ByVal YMin As Double, ByVal YMax As Double, ByVal ZMin As Double, ByVal ZMax As Double, Optional ByVal DeSelect As Boolean = False, Optional ByVal CSys As String = "Global", Optional ByVal IncludeIntersections As Boolean = False, Optional ByVal Point As Boolean = True, Optional ByVal Line As Boolean = True, Optional ByVal Area As Boolean = True, Optional ByVal Solid As Boolean = True, Optional ByVal Link As Boolean = True) As Long

## Parameters

XMin, XMax

The maximum and minimum X coordinates of the selection box in the specified coordinate system.

YMin, YMax

The maximum and minimum Y coordinates of the selection box in the specified coordinate system.

ZMin, ZMax

The maximum and minimum Z coordinates of the selection box in the specified coordinate system.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

CSys

The name of the coordinate system in which XMin, XMax, YMin, YMax, ZMin and ZMax are specified.

IncludeIntersections

When this item is True, objects that are inside the box or intersecting the sides and edges of the box are selected or deselected.

When this item is False, only objects that are fully inside the box are selected or deselected.

Point

Point objects that fall inside the box are only selected or deselected when this item is True.

Line

Line objects that fall inside the box are only selected or deselected when this item is True.

Area

Area objects that fall inside the box are only selected or deselected when this item is True.

Solid

Solid objects that fall inside the box are only selected or deselected when this item is True.

Link

Link objects that fall inside the box are only selected or deselected when this item is True.

## Remarks

This function selects or deselects objects inside the box defined by the XMin, XMax, YMin, YMax, ZMin and ZMax coordinates.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectCoordinateRange()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'select coordinate range

      ret = SapModel.SelectObj.CoordinateRange(-100, 100, 0, 0, 100, 200, , , True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
