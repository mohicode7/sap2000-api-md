# AddByCoord

## Syntax

SapObject.SapModel.CableObj.AddByCoord

## VB6 Procedure

Function AddByCoord(ByVal xi As Double, ByVal yi As Double, ByVal zi As Double, ByVal xj As Double, ByVal yj As Double, ByVal zj As Double, ByRef Name As String, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "", Optional ByVal CSys As String = "Global") As Long

## Parameters

xi, yi, zi

The coordinates of the I-End of the added cable object. The coordinates are in the coordinate system defined by the CSys item.

xj, yj, zj

The coordinates of the J-End of the added cable object. The coordinates are in the coordinate system defined by the CSys item.

Name

This is the name that the program ultimately assigns for the cable object. If no UserName is specified,n the program assigns a default name to the cable object. If a UserName is specified and that name is not used for another frame, cable or tendon object, the UserName is assigned to the cable object; otherwise a default name is assigned to the cable object.

PropName

This is Default or the name of a defined cable property.

If it is Default, the program assigns a default cable property to the cable object. If it is the name of a defined cable property, that property is assigned to the cable object.

UserName

This is an optional user specified name for the cable object. If a UserName is specified and that name is already used for another cable object, the program ignores the UserName.

CSys

The name of the coordinate system in which the cable object end point coordinates are defined.

## Remarks

This function adds a new cable object whose end points are at the specified coordinates.

The function returns zero if the cable object is successfully added, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddCableObjByCoord()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by coordinates

      ret = SapModel.CableObj.AddByCoord(-300, 0, 0, -100, 0, 124, Name)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByPoint](AddByPoint_%7BCable_Object%7D.md)

[SetCableData](SetCableData_%7BCable_Object%7D.md)
