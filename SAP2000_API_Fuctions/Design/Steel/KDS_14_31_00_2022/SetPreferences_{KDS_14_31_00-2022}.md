# SetPreference

## Syntax

SapObject.SapModel.DesignSteel.KDS143100\_2022.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 32, inclusive, indicating the preference item considered.

1 = Multi-response case design

2 = Framing type

3 = Seismic design category

4 = Importance Factor

5 = Design Systems Rho

6 = Design Systems Sds

7 = Design Systems R

8 = Design Systems Omega0

9 = Design Systems Cd

10 = Analysis method

11 = Second Order Method

12 = Stiffness Reduction Method

13 = Phi for bending

14 = Phi for compression

15 = Phi for tension yielding

16 = Phi for tension fracture

17 = Phi for shear

18 = Phi for shear short webbed rolled I

19 = Phi for torsion

20 = Ignore seismic code

21 = Ignore special seismic load

22 = Doubler plate is plug welded

23 = HSS welding type

24 = Reduce HSS thickness

25 = Consider deflection

26 = DL deflection limit, L/Value

27 = SDL + LL deflection limit, L/Value

28 = LL deflection limit, L/Value

29 = Total load deflection limit, L/Value

30 = Total camber limit, L/Value

31 = Pattern live load factor

32 = Demand/capacity ratio limit

Value

The value of the considered preference item.

1 = Multi-response case design

1 = Envelopes

2 = Step-by-step

3 = Last step

4 = Envelopes - All

5 = Step-by-step - All

2 = Framing type

1 = SMF

2 = IMF

3 = OMF

4 = SCBF

5 = OCBF

6 = OCBFI

7 = EBF

3 = Seismic design category

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

3 = Design provision

1 = LRFD

2 = ASD

4 = Importance Factor

Value > 0

5 = Design System Rho

Value > 0

6 = Design System Sds

Value >= 0

7 = Design System R

Value > 0

8 = Design System Omega0

Value > 0

9 = Design System Cd

Value > 0

10 = Analysis method

1 = Direct Analysis

2 = Effective Length

3 = Limited 1st Order Elastic

11 = Second Order Method

1 = General 2nd Order

2 = Amplified 1st Order

12 = Stiffness Reduction Method

1 = Tau-b variable

2 = Tau-b Fixed

3 = No Modification

13 = Phi or Omega for bending

Value > 0

14 = Phi or Omega for compression

Value > 0

15 = Phi or Omega for tension yielding

Value > 0

16 = Phi or Omega for tension fracture

Value > 0

17 = Phi or Omega for shear

Value > 0

18 = Phi or Omega for shear short webbed rolled I

Value > 0

19 = Phi or Omega for torsion

Value > 0

20 = Ignore seismic code

0 = No

Any other value = Yes

21 = Ignore special seismic load

0 = No

Any other value = Yes

22 = Doubler plate is plug welded

0 = No

Any other value = Yes

23 = HSS welding type

1 = ERW

2 = SAW

24 = Reduce HSS thickness

0 = No

Any other value = Yes

25 = Consider deflection

0 = No

Any other value = Yes

26 = DL deflection limit, L/Value

Value > 0

27 = SDL + LL deflection limit, L/Value

Value > 0

28 = LL deflection limit, L/Value

Value > 0

29 = Total load deflection limit, L/Value

Value > 0

30 = Total camber limit, L/Value

Value > 0

31 = Pattern live load factor

Value >= 0

32 = Demand/capacity ratio limit

Value > 0

## Remarks

This function sets the value of a steel design preference item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignPreferenceItemKDS143100_2022()

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

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



       'set steel design code

          ret = SapModel.DesignSteel.SetCode("KDS 14 31 00 : 2022")



       'set preference item

          ret = SapModel.DesignSteel.KDS143100_2022.SetPreference(1, 7)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version v27.1.0

## See Also

[GetPreference](GetPreferences_%7BKDS_14_31_00-2022%7D.md)
