# ChangeName

## Syntax

SapObject.SapModel.PropTendon.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined tendon property.

NewName

The new name for the tendon property.

## Remarks

This function changes the name of an existing tendon property.

The function returns zero if the new name is successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeTendonPropName()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_TENDON, , , , , , MATERIAL_TENDON_SUBTYPE_ASTM_A416Gr270)



   'set new tendon property

      ret = SapModel.PropTendon.SetProp("T1", Name, 1, 2.25)



   'change name of tendon property

      ret = SapModel.PropTendon.ChangeName("T1", "MyTendon")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
