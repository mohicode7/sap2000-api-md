# GetProperty

## Syntax

SapObject.SapModel.CableObj.GetProperty

## VB6 Procedure

Function GetProperty(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a defined cable object.

PropName

The name of the cable property assigned to the cable object.

## Remarks

This function retrieves the cable property assigned to a cable object.

The function returns zero if the cable object property is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableSectionProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropName As String

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



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'get cable property

      ret = SapModel.CableObj.GetProperty(Name, PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetProperty](SetProperty_%7BCable_Object%7D.md)
