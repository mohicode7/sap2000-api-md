# AddByCoord

## Syntax

SapObject.SapModel.LinkObj.AddByCoord

## VB6 Procedure

Function AddByCoord(ByVal xi As Double, ByVal yi As Double, ByVal zi As Double, ByVal xj As Double, ByVal yj As Double, ByVal zj As Double, ByRef Name As String, Optional IsSingleJoint As Boolean = False, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "", Optional ByVal CSys As String = "Global") As Long

## Parameters

xi, yi, zi

The coordinates of the I-End of the added link object. The coordinates are in the coordinate system defined by the CSys item.

xj, yj, zj

The coordinates of the J-End of the added link object. The coordinates are in the coordinate system defined by the CSys item.

These coordinates are ignored if the IsSingleJoint item is True.

Name

This is the name that the program ultimately assigns for the link object. If no UserName is specified, the program assigns a default name to the link object. If a UserName is specified and that name is not used for another link object, the UserName is assigned to the link object; otherwise a default name is assigned to the link object.

IsSingleJoint

This item is True if a one-joint link is added and False if a two-joint link is added.

PropName

This is either Default or the name of a defined link property.

If it is Default the program assigns a default link property to the link object. If it is the name of a defined link property, that property is assigned to the link object.

UserName

This is an optional user specified name for the link object. If a UserName is specified and that name is already used for another link object, the program ignores the UserName.

CSys

The name of the coordinate system in which the link object end point coordinates are defined.

## Remarks

This function adds a new link object whose end points are at the specified coordinates.

The function returns zero if the link object is successfully added; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddLinkObjByCoord()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name1 As String

      Dim Name2 As String



   'create Sap2000 object

      Set SapObject = New Sap200015.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add link object by coordinates

      ret = SapModel.LinkObj.AddByCoord(-288, 0, 288, 0, 0, 0, Name1, True)

      ret = SapModel.LinkObj.AddByCoord(-288, 0, 0, 0, 0, 144, Name2)



   'refresh view

      ret = SapModel.View.RefreshView



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByPoint](AddByPoint_%7BLink_Object%7D.md)
