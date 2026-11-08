# GetPreference

## Syntax

SapObject.SapModel.DesignConcrete.Eurocode\_2\_2023.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 28, inclusive, indicating the preference item considered.

1 = Country

2 = Combos equation

3 = Second order method

4 = Number of interaction curves

5 = Number of interaction points

6 = Consider minimum eccentricity

7 = Theta0

8 = Gamma steel

9 = Gamma concrete

14 = Pattern live load factor

15 = Utilization factor limit

16 = Multi-response case design

17 = Reliability Class

18 = Gamma concrete modulus

20 = Consider Torsion

21 = Longitudinal rebar size top

22 = Longitudinal rebar size bottom

23 = Is longitudinal rebar ribbed?

24 = Gamma shear

25 = kTC

26 = kTT

27 = Framing Type

28 = Ignore Beneficial Pu for Beam Design?

Value

The value of the considered preference item.

1 = Country

      1 = CEN Default

      2 = United Kingdom

      3 = Slovenia

      5 = Norway

      6 = Singapore

      7 = Sweden

      8 = Finland

      9 = Denmark

    10 = Portugal

    11 = Germany

    12 = Poland

    13 = Ireland

2 = Combos equation

      1 = Eq. 6.10

      2 = Max of Eqs. 6.10a and 6.10b

3 = Second order method

      1 = Nominal stiffness

      2 = Nominal curvature

      3 = None

4 = Number of interaction curves

Value >= 4 and divisible by 4

5 = Number of interaction points

Value >= 5 and odd

6 = Consider minimum eccentricity

0 = No

Any other value = Yes

7 = Theta0

Value > 0

8 = Gamma steel

Value > 0

9 = Gamma concrete

Value > 0

10 = AlphaCC

  Value > 0

11 = AlphaCT

  Value > 0

12 = AlphaLCC

  Value > 0

13 = AlphaLCT

  Value > 0

14 = Pattern live load factor

  Value >= 0

15 = Utilization factor limit

  Value > 0

16 = Multi-response case design

  1 = Envelopes

  2 = Step-by-step

  3 = Last step

  4 = Envelopes -- All

  5 = Step-by-step -- All

17 = Reliability Class

 1 = Class 1

 2 = Class 2

  3 = Class 3

                        18 = GammacE

                             Value > 0

19 = Alphae (not used)

                             Value > 0

20 = Consider torsion

0 = No

Any other value = Yes

21 = Longitudinal rebar size top

      Value is the index in the list of the rebar defined in the Reinforcing Bar Sizes form

22 = Longitudinal rebar size bottom      

      Value is the index in the list of the rebar defined in the Reinforcing Bar Sizes form

23 = Is longitudinal rebar ribbed?

0 = No

Any other value = Yes

24 = Gamma shear

  Value > 0

25 = kTC

  Value > 0

26 = kTT

  Value > 0

27 = Design for B/C capacity ratio

0 = No

Any other value = Yes

28 = Ignore Beneficial Pu for Beam Design?

0 = No

Any other value = Yes

## Remarks

This function retrieves the value of a concrete design preference item.

The function returns zero if the item is successfully retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignPreferenceItemEurocode_2_2023()

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

          ret = SapModel.InitializeNewModel



       'create new concrete frame section property

          ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



       'set concrete design code

          ret = SapModel.DesignConcrete.SetCode("Eurocode 2-2023")



       'get preference item

          ret = SapModel.DesignConcrete.Eurocode_2_2023.GetPreference(5, Value)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 27.0.0

## See Also

[SetPreference](SetPreference_%7BConcrete_Eurocode_2-2023%7D.md)
