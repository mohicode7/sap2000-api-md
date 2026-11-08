# SetShellDesign

## Syntax

SapObject.SapModel.PropArea.SetShellDesign

## VB6 Procedure

Function SetShellDesign(ByVal Name As String, ByVal MatProp As String, ByVal SteelLayoutOption As Long, ByVal DesignCoverTopDir1 As Double, ByVal DesignCoverTopDir2 As Double, ByVal DesignCoverBotDir1 As Double, ByVal DesignCoverBotDir2 As Double) As Long

## Parameters

Name

The name of an existing shell-type area property.

MatProp

The name of the material property for the area property.

SteelLayoutOption

This is 0, 1 or 2, indicating the rebar layout option.

0 = Default

1 = One layer

2 = Two layers

DesignCoverTopDir1, DesignCoverTopDir2

The cover to the centroid of the top reinforcing steel running in the local 1 and 2 axes directions of the area object, respectively. [L]

This item applies only when SteelLayoutOption = 1 or 2.

DesignCoverBotDir1, DesignCoverBotDir2

The cover to the centroid of the bottom reinforcing steel running in the local 1 and 2 axes directions of the area object, respectively. [L]

This item applies only when SteelLayoutOption = 2.

## Remarks

This function assigns the design parameters for shell-type area properties.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAreaPropShellDesign()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'set new area property

      ret = SapModel.PropArea.SetShell("A1", 1, "4000Psi", 0, 16, 16)



   'add ASTM A706 rebar material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)



   'set area property design parameters

      ret = SapModel.PropArea.SetShellDesign("A1", Name, 2, 2, 3, 2.5, 3.5)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetShellDesign](GetShellDesign.md)
