# ResetOverwrites

## Syntax

SapObject.SapModel.DesignColdFormed.ResetOverwrites

## VB6 Procedure

Function ResetOverwrites() As Long

## Parameters

None

## Remarks

This function resets all cold formed frame design overwrites to default values.

The function returns zero if the overwrites are successfully reset; otherwise it returns a nonzero value.

The function will fail if no cold formed frame objects are present.

## VBA Example



```
Sub ResetColdFormedDesignOverwrites()

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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", Name , 9, 3, 0.06, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 72, 2, 144, True, "CdC", "CdC")



   'reset cold formed design overwrites

      ret = SapModel.DesignColdFormed.ResetOverwrites



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also
