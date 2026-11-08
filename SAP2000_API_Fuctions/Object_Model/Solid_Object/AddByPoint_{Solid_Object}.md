# AddByPoint

## Syntax

SapObject.SapModel.SolidObj.AddByPoint

## VB6 Procedure

Function AddByPoint(ByRef Point() as String, ByRef Name As String, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "") As Long

## Parameters

Point

This is an array containing the names of the eight point objects that define the corner points of the added solid object.

Name

This is the name that the program ultimately assigns for the solid object. If no UserName is specified, the program assigns a default name to the solid object. If a UserName is specified and that name is not used for another solid object, the UserName is assigned to the solid object; otherwise a default name is assigned to the solid object.

PropName

This is either Default or the name of a defined solid property.

If it is Default, the program assigns a default solid property to the solid object. If it is the name of a defined solid property, that property is assigned to the solid object.

UserName

This is an optional user specified name for the solid object. If a UserName is specified and that name is already used for another solid object, the program ignores the UserName.

## Remarks

This function adds a new solid object whose corner points are specified by name.

The function returns zero if the solid object is successfully added; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddSolidObjByPoint()

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

      ret = SapModel.File.New3DFrame(Openframe, 2, 144, 3, 288, 2, 288)



   'add solid object by points

      Redim Point(7)

      Point(0) = "1"

      Point(1) = "10"

      Point(2) = "4"

      Point(3) = "13"

      Point(4) = "2"

      Point(5) = "11"

      Point(6) = "5"

      Point(7) = "14"

      ret = SapModel.SolidObj.AddByPoint(Point, Name)



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

[AddByCoord](AddByCoord_%7BSolid_Object%7D.md)
