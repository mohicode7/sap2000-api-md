# SetPreference

## Syntax

SapObject.SapModel.DesignSteel.AISC360\_10.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value
As Double) As Long

## Parameters

Item

This is an integer between 1 and 35, inclusive, indicating
the preference item considered.

1 = Framing type

2 = Seismic design category

3 = Design provision

4 = Analysis method (Obsolete,
replaced by 27, 28, and 29)

5 = Notional load coefficient

6 = Phi or Omega for bending

7 = Phi or Omega for compression

8 = Phi or Omega for tension
yielding

9 = Phi or Omega for tension
fracture

10 = Phi or Omega for
shear

11 = Phi or Omega for
shear short webbed rolled I

12 = Phi or Omega for
torsion

13 = Ignore seismic code

14 = Ignore special seismic
load

15 = Doubler plate is
plug welded

16 = HSS welding type

17 = Reduce HSS thickness

18 = Consider deflection

19 = DL deflection limit,
L/Value

20 = SDL + LL deflection
limit, L/Value

21 = LL deflection limit,
L/Value

22 = Total load deflection
limit, L/Value

23 = Total camber limit,
L/Value

24 = Pattern live load
factor

25 = Demand/capacity ratio
limit

26 = Multi-response case design  
27 = Analysis Method  
28 = Second Order Method  
29 = Stiffness Reduction Method

30 = Importance Factor

31 = Design Systems Rho

32 = Design Systems Sds

33 = Design Systems R

34 = Design Systems Omega0

35 = Design Systems Cd

Value

The value of the considered preference item.

1 = Framing type

1 = SMF

2 = IMF

3 = OMF

4 = SCBF

5 = OCBF

6 = OCBFI

7 = EBF

2 = Seismic design category

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

3 = Design provision

1 = LRFD

2 = ASD

4 = Analysis method (Obsolete,
replaced by 27, 28, and 29)

1 = Gen 2nd Order Elastic

2 = 2nd Order By Amp
1st Order

3 = Limited 1st Order
Elastic

4 = DAM Gen 2nd Order
Taub Variable

5 = DAM Gen 2nd Order
Taub Fixed

6 = DAM Amp 1st Order
Taub Variable

7 = DAM Amp 1st Order
Taub Fixed

5 = Notional load coefficient

Value > 0

6 = Phi or Omega for bending

Value > 0

7 = Phi or Omega for compression

Value > 0

8 = Phi or Omega for tension
yielding

Value > 0

9 = Phi or Omega for tension
fracture

Value > 0

10 = Phi or Omega for
shear

Value > 0

11 = Phi or Omega for
shear short webbed rolled I

Value > 0

12 = Phi or Omega for
torsion

Value > 0

13 = Ignore seismic code

0 = No

Any other value = Yes

14 = Ignore special seismic
load

0 = No

Any other value = Yes

15 = Doubler plate is
plug welded

0 = No

Any other value = Yes

16 = HSS welding type

1 = ERW

2 = SAW

17 = Reduce HSS thickness

0 = No

Any other value = Yes

18 = Consider deflection

0 = No

Any other value = Yes

19 = DL deflection limit,
L/Value

Value > 0

20 = SDL + LL deflection
limit, L/Value

Value > 0

21 = LL deflection limit,
L/Value

Value > 0

22 = Total load deflection
limit, L/Value

Value > 0

23 = Total camber limit,
L/Value

Value > 0

24 = Pattern live load
factor

Value >= 0

25 = Demand/capacity ratio
limit

Value > 0

26 = Multi-response case
design

  1 = Envelopes

  2 = Step-by-step

  3 = Last
step

  4 = Envelopes
-- All

  5 = Step-by-step
-- All

27 =  Analysis Method

1 = Direct Analysis

2 = Effective Length

3 = Limited 1st Order

28 =  Second Order
Method

1 = General 2nd Order

2 = Amplified 1st Order

29 =  Stiffness Reduction
Method

1 = Tau-b variable

2 = Tau-b Fixed

3 = No Modification

30 = Importance Factor

Value > 0

31 = Design System Rho

Value > 0

32 = Design System Sds

Value >= 0

33 = Design System R

Value > 0

34 = Design System Omega0

Value > 0

35 = Design System Cd

Value > 0

## Remarks

This function sets the value of a steel design preference
item.

The function returns zero if the item is successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignPreferenceItemAISC360_10()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("AISC360-10")



   'set preference item

      ret = SapModel.DesignSteel.AISC360_10.SetPreference(1, 
 7)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

Added items 30 through 35 in v17.3.0.

v23.2.0
- modified items 6 through 12 to allow phi or omega, depending on the
design provision value (item 3).

## See Also

[GetPreference](GetPreference_%7BSteel_AISC_360-10%7D.md)
