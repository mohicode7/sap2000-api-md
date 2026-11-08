# AddByCoord

## Syntax

SapObject.SapModel.SolidObj.AddByCoord

## VB6 Procedure

Function AddByCoord(ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, ByRef Name As String, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "", Optional ByVal CSys As String = "Global") As Long

## Parameters

x, y, z

These are arrays of x, y and z coordinates, respectively, for the corner points of the solid object. The coordinates are in the coordinate system defined by the CSys item.

Name

This is the name that the program ultimately assigns for the solid object. If no UserName is specified, the program assigns a default name to the solid object. If a UserName is specified and that name is not used for another solid object, the UserName is assigned to the solid object; otherwise a default name is assigned to the solid object.

PropName

This is either Default or the name of a defined solid property.

If it is Default, the program assigns a default solid property to the solid object. If it is the name of a defined solid property, that property is assigned to the solid object.

UserName

This is an optional user specified name for the solid object. If a UserName is specified and that name is already used for another solid object, the program ignores the UserName.

CSys

The name of the coordinate system in which the solid object point coordinates are defined.

## Remarks

This function adds a new solid object whose corner points are at the specified coordinates. Note that solid objects always are defined with eight corner points.

The function returns zero if the solid object is successfully added; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddSolidObjByCoord()

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



   'add solid object by coordinates

      ReDim x(7)

      ReDim y(7)

      ReDim z(7)

      x(0) = 0:    y(0) = 0:    z(0) = 0

      x(1) = 100:  y(1) = 0:    z(1) = 0

      x(2) = 0:    y(2) = 100:  z(2) = 0

      x(3) = 100:  y(3) = 100:  z(3) = 0

      x(4) = 0:    y(4) = 0:    z(4) = 100

      x(5) = 100:  y(5) = 0:    z(5) = 100

      x(6) = 0:    y(6) = 100:  z(6) = 100

      x(7) = 100:  y(7) = 100:  z(7) = 100

      ret = SapModel.SolidObj.AddByCoord(x, y, z, Name)



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

[AddByPoint](AddByPoint_%7BSolid_Object%7D.md)
