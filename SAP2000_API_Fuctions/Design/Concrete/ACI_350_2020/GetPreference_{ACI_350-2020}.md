# GetPreference

## Syntax

SapObject.SapModel.DesignConcreteShell.ACI350\_20.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 6, inclusive, indicating the preference item considered.

1 = Multi-response case design

2 = Consider environmental durability

3 = Consider shear design

4 = Shear design method

5 = Cotangent of the angle of concrete compressive strut

6 = Add tensile force due to shear

Value

The value of the considered preference item.

1 = Multi-response case design

1 = Envelopes

2 = Step-by step

3= Last step

4 = Envelopes - All

5 = Step-by step - All

2 = Consider environmental durability

0 = No

Any other value = Yes

3 = Consider shear design

0 = No

Any other value = Yes

4 = Shear design method

1 = Method 1 – considers increasing longitudinal reinforcement to increase concrete shear capacity up to the allowable limit. If insufficient, shear reinforcement will be added

2 = Method 2 – determines required shear reinforcement without considering to increase longitudinal reinforcement

5 = Cotangent of the angle of concrete compressive strut

Value > 0

6 = Add tensile force due to shear

0 = No

Any other value = Yes

## Remarks

This function retrieves the value of a concrete design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteShellDesignPreferenceItemACI350_20()
   'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim Value As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model
  
      ret = SapModel.InitializeNewModel(eUnits.kip_in_F)

   'create a wall model from template
      ret = SapModel.File.NewWall(6, 4, 6, 4)
       'set concrete shell design code

          ret = SapModel.DesignConcreteShell.SetCode("ACI 350-20")
   'set concrete shell design preference
      ret = SapModel.DesignConcreteShell.ACI350_20.SetPreference(5, 0.7)



       'get preference item

          ret = SapModel.DesignConcreteShell.ACI350_20.GetPreference(5, Value)



       'close Sap2000

          SapObject.ApplicationExit(False)

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  

```



## Release Notes

Initial release in version 26.0.0.

## See Also

[SetPreference](SetPreference_%7BACI_350-2020%7D.md)
