# GetComboAutoGenerate

## Syntax

SapObject.SapModel.DesignConcrete.GetComboAutoGenerate

## VB6 Procedure

Function GetComboAutoGenerate(ByRef AutoGenerate As Boolean) As Long

## Parameters

AutoGenerate

If this item is True, the option to automatically generate code-based design load combinations for concrete frame design is turned on. If it is False, the option is turned off.

## Remarks

This function retrieves the value of the automatically generated code-based design load combinations option for concrete frame design.

The function returns zero if the options are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignComboAutoGenerate()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim AutoGenerate As Boolean
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
   'set option to not auto generate code-based design load combinations

      ret = SapModel.DesignConcrete.GetComboAutoGenerate(AutoGenerate)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 18.1.0.

## See Also

[SetComboAutoGenerate](SetComboAutoGenerate_%7BConcrete%7D.md)
