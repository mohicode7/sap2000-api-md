# GetPreference

## Syntax

SapObject.SapModel.DesignAluminum.AA\_LRFD\_2000.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 6, inclusive, indicating the preference item considered.

1 = Framing type

2 = Demand/capacity ratio limit

3 = Phiy resistance factor

4 = Phib resistance factor

5 = Phic resistance factor

6 = Phiu resistance factor

7 = Phicc resistance factor

8 = Phicp resistance factor

9 = Phiv resistance factor

10 = Phivp resistance factor

11 = Phiw resistance factor

12 = Time history design

Value

The value of the considered preference item.

1 = Framing type

1 = Moment Frame

2 = Braced Frame

2 = Demand/capacity ratio limit

Value > 0

3 = Phiy resistance factor

Value > 0

4 = Phib resistance factor

Value > 0

5 = Phic resistance factor

Value > 0

6 = Phiu resistance factor

Value > 0

7 = Phicc resistance factor

Value > 0

8 = Phicp resistance factor

Value > 0

9 = Phiv resistance factor

Value > 0

10 = Phivp resistance factor

Value > 0

11 = Phiw resistance factor

Value > 0

12 = Time history design

1 = Envelopes

2 = Step-by step

## Remarks

This function retrieves the value of an aluminum design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignPreferenceItemAA_LRFD_2000()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim Value As Double



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

      ret = SapModel.DesignAluminum.SetCode("AA-LRFD 2000")



   'get preference item

      ret = SapModel.DesignAluminum.AA_LRFD_2000.GetPreference(1, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetPreference](SetPreference_%7BAluminum_AA_LRFD_2000%7D.md)
