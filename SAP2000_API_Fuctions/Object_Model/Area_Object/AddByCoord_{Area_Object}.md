# AddByCoord

## Syntax

SapObject.SapModel.AreaObj.AddByCoord

## VB6 Procedure

Function AddByCoord(ByVal NumberPoints As long, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, ByRef Name As String, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "", Optional ByVal CSys As String = "Global") As Long

## Parameters

NumberPoints

The number of points in the area abject.

x, y, z

These are arrays of x, y and z coordinates, respectively, for the corner points of the area object. The coordinates are in the coordinate system defined by the CSys item. The coordinates should be ordered to run clockwise or counter clockwise around the area object.

Name

This is the name that the program ultimately assigns to the area object. If no UserName is specified, the program assigns a default name to the area object. If a UserName is specified and that name is not used for another area object, the UserName is assigned to the area object; otherwise a default name is assigned to the area object.

PropName

This is Default, None or the name of a defined area property.

If it is Default, the program assigns a default area property to the area object. If it is None, no area property is assigned to the area object. If it is the name of a defined area property, that property is assigned to the area object.

UserName

This is an optional user specified name for the area object. If a UserName is specified and that name is already used for another area object, the program ignores the UserName.

CSys

The name of the coordinate system in which the area object point coordinates are defined.

## Remarks

This function adds a new area object, defining points at the specified coordinates.

The function returns zero if the area object is successfully added, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddAreaObjByCoord()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double

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

      ret = SapModel.File.NewBlank



   'add area object by coordinates

      ReDim x(5)

      ReDim y(5)

      ReDim z(5)

      x(0) = 50:   y(0) = 0

      x(1) = 100:  y(1) = 0

      x(2) = 150:  y(2) = 40

      x(3) = 100:  y(3) = 80

      x(4) = 50:   y(4) = 80

      x(5) = 0:    y(5) = 40

      ret = SapModel.AreaObj.AddByCoord(6, x, y, z, Name)



   'refresh view

      ret = SapModel.View.RefreshView(0, False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByPoint](AddByPoint_%7BArea_Object%7D.md)
