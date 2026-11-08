# ResetOverwrites

## Syntax

SapObject.SapModel.DesignAluminum.ResetOverwrites

## VB6 Procedure

Function ResetOverwrites() As Long

## Parameters

None

## Remarks

This function resets all aluminum frame design overwrites to default values.

The function returns zero if the overwrites are successfully reset; otherwise it returns a nonzero value.

The function will fail if no aluminum frame objects are present.

## VBA Example



```
Sub ResetAluminumDesignOverwrites()

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



   'reset aluminum design overwrites

      ret = SapModel.DesignAluminum.ResetOverwrites



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also
