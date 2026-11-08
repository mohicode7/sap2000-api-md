# GetConnectivity

## Syntax

SapObject.SapModel.PointObj.GetConnectivity

## VB6 Procedure

Function GetConnectivity(ByVal Name As String, ByRef NumberItems As Long, ByRef ObjectType() As Long, ByRef ObjectName() As String, ByRef PointNumber() As Long) As Long

## Parameters

Name

The name of an existing point object.

NumberItems

This is the total number of objects connected to the specified point object.

ObjectType

This is an array that includes the object type of each object connected to the specified point object.

2 = Frame object

3 = Cable object

4 = Tendon object

5 = Area object

6 = Solid object

7 = Link object

ObjectName

This is an array that includes the object name of each object connected to the specified point object.

PointNumber

This is an array that includes the point number within the considered object that corresponds to the specified point object.

## Remarks

This function returns a list of objects connected to a specified point object.

The function returns zero if the list is successfully filled; otherwise it returns nonzero.

## VBA Example



```
Sub GetPointObjectConnectivity()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems as Long

      Dim ObjectType() As Long

      Dim ObjectName() As String

      Dim PointNumber() As Long



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



   'get objects connected to point object 11

      ret = SapModel.PointObj.GetConnectivity("11", NumberItems, ObjectType, ObjectName, PointNumber)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also
