# AddByCoord

## Syntax

SapObject.SapModel.TendonObj.AddByCoord

## VB6 Procedure

Function AddByCoord(ByVal xi As Double, ByVal yi As Double, ByVal zi As Double, ByVal xj As Double, ByVal yj As Double, ByVal zj As Double, ByRef Name As String, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "", Optional ByVal CSys As String = "Global") As Long

## Parameters

xi, yi, zi

The coordinates of the I-End of the added tendon object. The coordinates are in the coordinate system defined by the CSys item.

xj, yj, zj

The coordinates of the J-End of the added tendon object. The coordinates are in the coordinate system defined by the CSys item.

Name

This is the name that the program ultimately assigns for the tendon object. If no UserName is specified, the program assigns a default name to the tendon object. If a UserName is specified and that name is not used for another frame, cable or tendon object, the UserName is assigned to the tendon object; otherwise a default name is assigned to the tendon object.

PropName

This is Default, None or the name of a defined tendon property.

If it is Default, the program assigns a default tendon property to the tendon object. If it is None, no tendon property is assigned to the tendon object. If it is the name of a defined tendon property, that property is assigned to the tendon object.

UserName

This is an optional user specified name for the tendon object. If a UserName is specified and that name is already used for another tendon object, the program ignores the UserName.

CSys

The name of the coordinate system in which the tendon object end point coordinates are defined.

## Remarks

This function adds a new tendon object whose end points are at the specified coordinates.

The function returns zero if the tendon object is successfully added, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddTendonObjByCoord()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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



   'add tendon object by coordinates

      ret = SapModel.TendonObj.AddByCoord(-288, 0, 288, 288, 0, 288, Name)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByPoint](AddByPoint_%7BTendon_Object%7D.md)

[SetTendonData](SetTendonData.md)
