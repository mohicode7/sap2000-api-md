# GetSelected

## Syntax

SapObject.SapModel.SelectObj.GetSelected

## VB6 Procedure

Function GetSelected(ByRef NumberItems As Long, ByRef ObjectType() As Long, ByRef ObjectName() As String) As Long

## Parameters

NumberItems

The number of selected objects.

ObjectType

This is an array that includes the object type of each selected object.

1 = Point object

2 = Frame object

3 = Cable object

4 = Tendon object

5 = Area object

6 = Solid object

7 = Link object

ObjectName

This is an array that includes the name of each selected object.

## Remarks

This function retrieves a list of selected objects.

The function returns zero if the selection list is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSelectedObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim ObjectType() As Long

      Dim ObjectName() As String



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



   'set all point objects selected

      ret = SapModel.PointObj.SetSelected("ALL", True, Group)



   'set all frame objects selected

      ret = SapModel.FrameObj.SetSelected("ALL", True, Group)



   'get selected objects

      ret = SapModel.SelectObj.GetSelected(NumberItems, ObjectType, ObjectName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
