# GetNameList

## Syntax

SapObject.SapModel.PropTendon.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of tendon property names retrieved by the program.

MyName

This is a one-dimensional array of tendon property names. The MyName array is created as a dynamic, zero-based, array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the SAP2000 program, filled with the names, and returned to the API user.

## Remarks

This function retrieves the names of all defined tendon properties in the model.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetTendonNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberNames As Long

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



   'add tendon material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_TENDON, , , , , , MATERIAL_TENDON_SUBTYPE_ASTM_A416Gr270)



   'set new tendon properties

      ret = SapModel.PropTendon.SetProp("T1", Name, 1, 2.25)

      ret = SapModel.PropTendon.SetProp("T2", Name, 1, 3.25)



   'get tendon property names

      ret = SapModel.PropTendon.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
