# GetPreference

## Syntax

SapObject.SapModel.DesignAluminum.AA\_2020.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 15, inclusive, indicating the preference item considered.

1 = Demand/capacity ratio limit

2 = Design Provision

3 = Lateral Factor

4 = Use Lateral Factor

5 = PhiTy or OmegaTy safety factor

6 = PhiTr or OmegaTr safety factor

7 = PhiC or OmegaC safety factor

9 = PhiBo or OmegaBo safetyfactor

9 = PhiBr or OmegaBr safetyfactor

10 = PhicVo or OmegaVo safetyfactor

11 = PhiVr or OmegaVr safetyfactor

12 = Time history design

13 = Analysis Method

14 = Second Order Method

15 = Stiffness Reduction Method

Value

The value of the considered preference item.

1 = Demand/capacity ratio limit

Value > 0

2 = Design Provision

1 = LRFD

2 = ASD

3 = Lateral Factor

Value > 0

4 = Use Lateral Factor

1 = No

2 = Yes

5 = PhiTy or OmegaTy safety factor

PhiTy:           0 < Value ≤ 1

OmegaTy:     1 ≤ Value

6 = PhiTr or OmegaTr safety factor

PhiTr:           0 < Value ≤ 1

OmegaTr:     1 ≤ Value

7 = PhiC or OmegaC safety factor

PhiC:           0 < Value ≤ 1

OmegaC:     1 ≤ Value

8 = PhiBo or OmegaBo safetyfactor

PhiBo:           0 < Value ≤ 1

OmegaBo:     1 ≤ Value

9 = PhiBr or OmegaBr safetyfactor

 PhiBr:           0 < Value ≤ 1

 OmegaBr:     1 ≤ Value

10 = PhicVo or OmegaVo safetyfactor

 PhiVo:           0 < Value ≤ 1

 OmegaVo:     1 ≤ Value

11 = PhiVr or OmegaVr safetyfactor

 PhiVr:           0 < Value ≤ 1

 OmegaVr:     1 ≤ Value

12 = Time history design

1 = Envelopes

2 = Step-by-step

13 = Analysis method

1 = Direct Analysis Method

2 = Effective Length

3 = Limited 1st Order

14 = Second Order Method

1 = General 2nd Order

2 = Amplified 1st Order

15 = Stiffness Reduction method

1 = Tau-b Variable

2 = Tau-b Fixed

3 = No Modification

If Design Provision is 1 (LRFD), the values of items 5through 11 will come from the Phi factors. Otherwise, if Design Provision is 2 (ASD) , the values of items 5 through 11 will come from the Omega factors .

## Remarks

This function retrieves the value of an aluminum design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignPreferenceItemAA_2020()

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

          ret = SapModel.DesignAluminum.SetCode("AA 2020")



       'get preference item

          ret = SapModel.DesignAluminum.AA_2020.GetPreference(1, Value)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 23.2.0

Added items 13 through 15 in version 27.0.0

## See Also

[SetPreference](SetPreference_%7BAluminum_AA_2020%7D.md)
