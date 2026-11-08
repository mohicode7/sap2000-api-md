# GetConnectivity

## Syntax

SapObject.SapModel.PointElm.GetConnectivity

## VB6 Procedure

Function GetConnectivity(ByVal Name As String, ByRef NumberItems As Long, ByRef ObjectType() As Long, ByRef ObjectName() As String, ByRef PointNumber() As Long) As Long

## Parameters

Name

The name of an existing point element.

NumberItems

This is the total number of elements connected to the specified point element.

ObjectType

This is an array that includes the element type of each element connected to the specified point element.

2 = Frame element

3 = Cable element

4 = Tendon element

5 = Area element

6 = Solid element

7 = Link element

ObjectName

This is an array that includes the element name of each element connected to the specified point element.

PointNumber

This is an array that includes the point number within the considered element that corresponds to the specified point element.

## Remarks

This function returns a list of elements connected to a specified point element.

The function returns zero if the list is successfully filled; otherwise it returns nonzero.

## VBA Example



```
Sub GetPointElementConnectivity()

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



       'create analysis model

          ret = SapModel.Analyze.CreateAnalysisModel



       'get elements connected to point element 11

          ret = SapModel.PointElm.GetConnectivity("11", NumberItems, ObjectType, ObjectName, PointNumber)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 12.00.

## See Also
