# AddDesignDefaultCombos

## Syntax

SapObject.SapModel.RespCombo.AddDesignDefaultCombos

## VB6 Procedure

Function AddDesignDefaultCombos(ByVal DesignSteel As Boolean, ByVal DesignConcrete As Boolean, ByVal DesignAluminum As Boolean, ByVal DesignColdFormed As Boolean) As Long

## Parameters

DesignSteel

If this item is True, default steel design combinations are to be added to the model.

DesignConcrete

If this item is True, default concrete design combinations are to be added to the model.

DesignAluminum

If this item is True, default aluminum design combinations are to be added to the model.

DesignColdFormed

If this item is True, default cold formed design combinations are to be added to the model.

## Remarks

This function adds a new default design load combinations to the model.

The function returns zero if the load combinations are successfully added, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddNewDefaultSteelDesignCombos()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'add combo

      ret = SapModel.RespCombo.AddDesignDefaultCombos(True, False, False, False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[Add](Add_%7BCombinations%7D.md)
