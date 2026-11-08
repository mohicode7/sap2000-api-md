# Add

## Syntax

SapObject.SapModel.RespCombo.Add

## VB6 Procedure

Function Add(ByVal Name As String, ByVal ComboType As Long) As Long

## Parameters

Name

The name of a new load combination.

ComboType

This is 0, 1, 2, 3 or 4 indicating the load combination type.

0 = Linear Additive

1 = Envelope

2 = Absolute Additive

3 = SRSS

4 = Range Additive

## Remarks

This function adds a new load combination.

The function returns zero if the load combination is successfully added, otherwise it returns a nonzero value.

The new load combination must have a different name from all other load combinations and all load cases. If the name is not unique, an error will be returned.

## VBA Example



```
Sub AddNewCombo()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add combo

      ret = SapModel.RespCombo.Add("COMB1", 1)



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

[AddDesignDefaultCombos](AddDesignDefaultCombos.md)
