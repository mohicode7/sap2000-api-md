# SetPreference

## Syntax

SapObject.SapModel.DesignAluminum.AA\_ASD\_2000.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 6, inclusive, indicating the preference item considered.

1 = Framing type

2 = Demand/capacity ratio limit

3 = Lateral factor

4 = Use lateral factor

5 = Bridge type structure

6 = Time history design

Value

The value of the considered preference item.

1 = Framing type

1 = Moment Frame

2 = Braced Frame

2 = Demand/capacity ratio limit

Value > 0

3 = Lateral factor

Value > 0

4 = Use lateral factor

0 = No

Any other value = Yes

5 = Bridge type structure

0 = No

Any other value = Yes

6 = Time history design

1 = Envelopes

2 = Step-by step

## Remarks

This function sets the value of an aluminum design preference item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAluminumDesignPreferenceItemAA_ASD_2000()

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



   'set aluminum design code

      ret = SapModel.DesignAluminum.SetCode("AA-ASD 2000")



   'set preference item

      ret = SapModel.DesignAluminum.AA_ASD_2000.SetPreference(1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetPreference](GetPreference_%7BAluminum_AA_ASD_2000%7D.md)
