# AddByPoint

## Syntax

SapObject.SapModel.AreaObj.AddByPoint

## VB6 Procedure

Function AddByPoint(ByVal NumberPoints as Long, ByRef Point() as String, ByRef Name As String, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "") As Long

## Parameters

NumberPoints

The number of points in the area abject.

Point

This is an array containing the names of the point objects that define the added area object. The point object names should be ordered to run clockwise or counter clockwise around the area object.

Name

This is the name that the program ultimately assigns for the area object. If no UserName is specified, the program assigns a default name to the area object. If a UserName is specified and that name is not used for another area object, the UserName is assigned to the area object; otherwise a default name is assigned to the area object.

PropName

This is Default, None or the name of a defined area property.

If it is Default, the program assigns a default area property to the area object. If it is None, no area property is assigned to the area object. If it is the name of a defined area property, that property is assigned to the area object.

UserName

This is an optional user specified name for the area object. If a UserName is specified and that name is already used for another area object, the program ignores the UserName.

## Remarks

This function adds a new area object whose defining points are specified by name.

The function returns zero if the area object is successfully added; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddAreaObjByPoint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Point() As String

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add area object by points

      Redim Point(3)

      Point(0) = "1"

      Point(1) = "4"

      Point(2) = "5"

      Point(3) = "2"

      ret = SapModel.AreaObj.AddByPoint(4, Point, Name)



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

[AddByCoord](AddByCoord_%7BArea_Object%7D.md)
