# SetGroup

## Syntax

SapObject.SapModel.DesignAluminum.SetGroup

## VB6 Procedure

Function SetGroup(ByVal Name As String, ByVal Selected As Boolean) As Long

## Parameters

Name

The name of an existing group.

Selected

If this item is True, the specified group is selected as a design group for aluminum design. If it is False, the group is not selected for aluminum design.

## Remarks

This function selects or deselects a group for aluminum design.

The function returns zero if the selection status is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAluminumDesignGroup()

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



   'add aluminum material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_ALUMINUM, , , MATERIAL_ALUMINUM_SUBTYPE_6061_T6)



   'create new aluminum frame section property

      ret = SapModel.PropFrame.SetISection("AI", Name , 18, 6, 0.5, 0.3, 6, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "AI", "AI")



   'select group for aluminum design

      ret = SapModel.DesignAluminum.SetGroup("ALL", True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetGroup](GetGroup_%7BAluminum%7D.md)
