# GetGroup

## Syntax

SapObject.SapModel.DesignSteel.GetGroup

## VB6 Procedure

Function GetGroup(ByRef NumberItems As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of groups selected for steel design.

MyName

This is an array that includes the name of each group selected for steel design.

## Remarks

This function retrieves the names of all groups selected for steel design.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignGroup()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim MyName() As String



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



   'select group for steel design

      ret = SapModel.DesignSteel.SetGroup("ALL", True)



   'get groups selected for steel design

      ret = SapModel.DesignSteel.GetGroup(NumberItems, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetGroup](SetGroup_%7BSteel%7D.md)
