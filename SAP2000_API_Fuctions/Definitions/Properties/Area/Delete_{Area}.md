# Delete

## Syntax

SapObject.SapModel.PropArea.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing area property.

## Remarks

The function deletes a specified area property.

The function returns zero if the area property is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified area property can not be deleted, for example, if it is being used by an area object.

## VBA Example



```
Sub DeleteAreaProp()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'set new area property

      ret = SapModel.PropArea.SetShell("A1", 1, "4000Psi", 0, 16, 16)



   'get area object names

      ret = SapModel.AreaObj.GetNameList(NumberNames, MyName)



   'assign area property

      For i = 1 to NumberNames

         ret = SapModel.AreaObj.SetProperty(MyName(i - 1), "A1")

      Next i



   'delete area property

      ret = SapModel.PropArea.Delete("ASEC1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
