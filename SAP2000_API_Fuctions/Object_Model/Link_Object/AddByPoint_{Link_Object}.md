# AddByPoint

## Syntax

SapObject.SapModel.LinkObj.AddByPoint

## VB6 Procedure

Function AddByPoint(ByVal Point1 as String, ByVal Point2 as String, ByRef Name As String, Optional IsSingleJoint As Boolean = False, Optional ByVal PropName As String = "Default", Optional ByVal UserName As String = "") As Long

## Parameters

Point1

The name of a defined point object at the I-End of the added link object.

Point2

The name of a defined point object at the J-End of the added link object.

This item is ignored if the IsSingleJoint item is True.

Name

This is the name that the program ultimately assigns for the link object. If no UserName is specified, the program assigns a default name to the link object. If a UserName is specified and that name is not used for another link object, the UserName is assigned to the link object; otherwise a default name is assigned to the link object.

IsSingleJoint

This item is True if a one-joint link is added and False if a two-joint link is added.

PropName

This is either Default or the name of a defined link property.

If it is Default the program assigns a default link property to the link object. If it is the name of a defined link property, that property is assigned to the link object.

UserName

This is an optional user specified name for the link object. If a UserName is specified and that name is already used for another link object, the program ignores the UserName.

## Remarks

This function adds a new link object whose end points are specified by name.

The function returns zero if the link object is successfully added; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddLinkObjByPoint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name1 As String

      Dim Name2 As String



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



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("3", "", Name1, True)

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name2)



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

[AddByCoord](AddByCoord_%7BLink_Object%7D.md)
