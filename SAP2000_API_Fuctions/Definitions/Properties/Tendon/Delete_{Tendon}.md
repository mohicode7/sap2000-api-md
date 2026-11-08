# Delete

## Syntax

SapObject.SapModel.PropTendon.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing tendon property.

## Remarks

The function deletes a specified tendon property.

The function returns zero if the property is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified property can not be deleted, for example, if it is assigned to an existing object.

## VBA Example



```
Sub DeleteTendonProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_TENDON, , , , , , MATERIAL_TENDON_SUBTYPE_ASTM_A416Gr270)



   'set new tendon properties

      ret = SapModel.PropTendon.SetProp("T1", Name, 1, 2.25)

      ret = SapModel.PropTendon.SetProp("T2", Name, 1, 3.25)



   'delete tendon property

      ret = SapModel.PropTendon.Delete("T1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
