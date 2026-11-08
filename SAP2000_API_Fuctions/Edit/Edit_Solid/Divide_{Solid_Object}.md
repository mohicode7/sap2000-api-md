# Divide

## Syntax

SapObject.SapModel.EditSolid.Divide

## VB6 Procedure

Function Divide(ByVal Name As String, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByRef NumberSolids As Long, ByRef SolidName() As String) As Long

## Parameters

Name

The name of an existing solid object.

n1

This is the number of objects created between faces 2 and 4 of the solid object.

n2

This is the number of objects created between faces 1 and 3 of the solid object.

n3

This is the number of objects created between faces 5 and 6 of the solid object.

NumberSolids

The number of solid objects created when the specified solid object is divided.

SolidName

This is an array of the name of each solid object created when the specified solid object is divided.

## Remarks

This function meshes solid objects.

The function returns zero if the meshing is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub DivideSolidObject()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberSolids As Long

      Dim SolidName() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

         ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'divide solid object

      ret = SapModel.EditSolid.Divide("1", 2, 3, 4, NumberSolids, SolidName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
