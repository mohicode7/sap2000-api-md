# AddCylindrical

## Syntax

SapObject.SapModel.PointObj.AddCylindrical

## VB6 Procedure

Function AddCylindrical(ByVal r As Double, ByVal Theta As Double, ByVal z As Double, ByRef Name As String, Optional ByVal userName As String = "", Optional ByVal csys As String = "Global", Optional ByVal MergeOff As Boolean = False, Optional ByVal MergeNumber As Long = 0) As Long

## Parameters

r

The radius for the added point object in the specified coordinate system. [L]

Theta

The angle for the added point object in the specified coordinate system. The angle is measured in the XY plane from the positive global X axis. When looking in the XY plane with the positive Z axis pointing toward you, a positive Theta angle is counter clockwise. [deg]

z

The Z-coordinate of the added point object in the specified coordinate system. [L]

Name

This is the name that the program ultimately assigns for the point object. If no UserName is specified, the program assigns a default name to the point object. If a UserName is specified and that name is not used for another point, the UserName is assigned to the point; otherwise a default name is assigned to the point.

If a point is merged with another point, this will be the name of the point object with which it was merged.

UserName

This is an optional user specified name for the point object. If a UserName is specified and that name is already used for another point object, the program ignores the UserName.

CSys

The name of the coordinate system in which the joint coordinates are defined.

MergeOff

If this item is False, a new point object that is added at the same location as an existing point object will be merged with the existing point object (assuming the two point objects have the same MergeNumber) and thus only one point object will exist at the location.

If this item is True, the points will not merge and two point objects will exist at the same location.

MergeNumber

Two points objects in the same location will merge only if their merge number assignments are the same. By default all pointobjects have a merge number of zero.

## Remarks

This function adds a point object to a model. The added point object will be tagged as a Special Point except if it was merged with another point object. Special points are allowed to exist in the model with no objects connected to them

The function returns zero if the point object is successfully added or merged, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddPointCylindrical()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim r As Double, Theta As Double, z As Double

      Dim Name as String

      Dim MyName as String

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model blank from template

      ret = SapModel.File.NewBlank



   'add point object to model

      r = 12

      Theta = 37

      z = 0

      MyName = "A1"

      ret = SapModel.PointObj.AddCartesian(r, Theta, z, Name, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
 
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddCartesian](AddCartesian.md)

[AddSpherical](AddSpherical.md)

[GetCoordCylindrical](GetCoordCylindrical_%7BPoint_Object%7D.md)

[GetSpecialPoint](GetSpecialPoint.md)

[SetSpecialPoint](SetSpecialPoint.md)
