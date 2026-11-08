# SetProp

## Syntax

SapObject.SapModel.PropCable.SetProp

## VB6 Procedure

Function SetProp(ByVal Name As String, ByVal MatProp As String, ByVal Area As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new cable property. If this is an existing property, that property is modified; otherwise, a new property is added.

MatProp

The name of the material property assigned to the cable property.

Area

The cross-sectional area of the cable. [L2]

Color

The display color assigned to the property. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, the program assigns a GUID to the property.

## Remarks

This function defines a cable property.

The function returns zero if the property is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCableProperty()

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



   'add material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_TENDON, , , , , , MATERIAL_TENDON_SUBTYPE_ASTM_A416Gr270)



   'set new cable property

      ret = SapModel.PropCable.SetProp("C1", Name, 2.25)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetProp](GetProp_%7BCable%7D.md)
