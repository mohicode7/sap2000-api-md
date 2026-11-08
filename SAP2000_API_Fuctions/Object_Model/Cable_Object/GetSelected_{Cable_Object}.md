# GetSelected

## Syntax

Sap2000.CableObj.GetSelected

## VB6 Procedure

Function GetSelected(ByVal Name As String, ByRef Selected As Boolean) As Long

## Parameters

Name

The name of an existing cable object.

Selected

This item returns True if the specified cable object is selected, otherwise it returns False.

## Remarks

This function retrieves the selected status for a cable object.

The function returns zero if the selected status is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableObjectSelectedStatus()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Selected As Boolean

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name1)

      ret = SapModel.CableObj.AddByPoint("5", "10", Name2)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name1, 7, 1, 0, 0, 24)

      ret = SapModel.CableObj.SetCableData(Name2, 7, 1, 0, 0, 24)



   'set all cables selected

      ret = SapModel.CableObj.SetSelected("All", True, Group)



   'get cable object selected status

      ret = SapModel.CableObj.GetSelected(Name1, Selected)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSelected](SetSelected_%7BCable_Object%7D.md)
