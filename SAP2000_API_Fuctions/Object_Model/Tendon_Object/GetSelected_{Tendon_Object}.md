# GetSelected

## Syntax

Sap2000.TendonObj.GetSelected

## VB6 Procedure

Function GetSelected(ByVal Name As String, ByRef Selected As Boolean) As Long

## Parameters

Name

The name of an existing tendon object.

Selected

This item returns True if the specified tendon object is selected, otherwise it returns False.

## Remarks

This function retrieves the selected status for a tendon object.

The function returns zero if the selected status is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonObjectSelectedStatus()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name1)

      ret = SapModel.TendonObj.AddByPoint("2", "8", Name2)



   'set all tendons selected

      ret = SapModel.TendonObj.SetSelected("All", True, Group)



   'get tendon object selected status

      ret = SapModel.TendonObj.GetSelected(Name1, Selected)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSelected](SetSelected_Tendon_Object%7D.md)
