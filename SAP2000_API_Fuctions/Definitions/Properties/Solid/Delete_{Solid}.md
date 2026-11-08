# Delete

## Syntax

SapObject.SapModel.PropSolid.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing solid property.

## Remarks

The function deletes a specified solid property.

The function returns zero if the property is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified property can not be deleted, for example, if it is assigned to an existing object.

## VBA Example



```
Sub DeleteSolidProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

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

      ret = SapModel.File.NewSolidBlock(20, 50, 20)



   'set new solid property

      ret = SapModel.PropSolid.SetProp("S1", "4000Psi", 0, 10, 20, True)



   'get solid object names

      ret = SapModel.SolidObj.GetNameList(NumberNames, MyName)



   'set solid property

      For i = 1 to NumberNames

         ret = SapModel.SolidObj.SetProperty(MyName(i - 1), "S1")

      Next i



   'delete solid property

      ret = SapModel.PropSolid.Delete("SOLID1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
