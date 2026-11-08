# SetPreference

## Syntax

SapObject.SapModel.DesignSteel.API\_4F\_2020.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 32, inclusive, indicating the preference item considered.

1 = Framing type

2 = Seismic design category

3 = Importance Factor

4 = Design Systems Rho

5 = Design Systems Sds

6 = Design Systems R

7 = Design Systems Omega0

8 = Design Systems Cd

9 = Analysis Method

10 = Second Order Method

11 = Stiffness Reduction Method

12 = Omega for bending

13 = Omega for compression

14 = Omega for tension yielding

15 = Omega for tension fracture

16 = Omega for shear

17 = Omega for shear short webbed rolled I

18 = Omega for torsion

19 = Ignore seismic code

20 = Ignore special seismic load

21 = Doubler plate is plug welded

22 = HSS welding type

23 = Reduce HSS thickness

24 = Consider deflection

25 = DL deflection limit, L/Value

26 = SDL + LL deflection limit, L/Value

27 = LL deflection limit, L/Value

28 = Total load deflection limit, L/Value

29 = Total camber limit, L/Value

30 = Pattern live load factor

31 = Demand/capacity ratio limit

32 =Multi-response case design

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

3 = Importance Factor

Value > 0

4 = Design System Rho

Value > 0

5 = Design System Sds

Value >= 0

6 = Design System R

Value > 0

7 = Design System Omega0

Value > 0

8 = Design System Cd

Value > 0

9 = Analysis Method

1 = Direct Analysis

2 = Effective Length

3 = Limited 1st Order

10 = Second Order Method

1 = General 2nd Order

2 = Amplified 1st Order

11 = Stiffness Reduction Method

1 = Tau-b variable

2 = Tau-b Fixed

3 = No Modification

12 = Omega for bending

Value > 0

13 = Omega for compression

Value > 0

14 = Omega for tension yielding

Value > 0

15 = Omega for tension fracture

Value > 0

16 = Omega for shear

Value > 0

17 = Omega for shear short webbed rolled I

Value > 0

18 = Omega for torsion

Value > 0

19 = Ignore seismic code

0 = No

Any other value = Yes

20 = Ignore special seismic load

0 = No

Any other value = Yes

21 = Doubler plate is plug welded

0 = No

Any other value = Yes

22 = HSS welding type

1 = ERW

2 = SAW

23 = Reduce HSS thickness

0 = No

Any other value = Yes

24 = Consider deflection

0 = No

Any other value = Yes

25 = DL deflection limit, L/Value

Value > 0

26 = SDL + LL deflection limit, L/Value

Value > 0

27 = LL deflection limit, L/Value

Value > 0

28 = Total load deflection limit, L/Value

Value > 0

29 = Total camber limit, L/Value

Value > 0

30 = Pattern live load factor

Value >= 0

31 = Demand/capacity ratio limit

Value > 0

32 = Multi-response case design

 1 = Envelopes

 2 = Step-by-step

 3 = Last step

 4 = Envelopes -- All

 5 = Step-by-step -- All

## Remarks

This function sets the value of a steel design preference item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignPreferenceItemAPI_4F_2020()

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

    ret = SapModel.DesignSteel.SetCode("API 4F-2020")



    'set preference item

    ret = SapModel.DesignSteel.API_4F_2020.SetPreference(1, 7)



    'close Sap2000

    SapObject.ApplicationExit False

    Set SapModel = Nothing

    Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 25.2.0

## See Also

[GetPreference](GetPreferences_%7BAPI_4F_2020%7D.md)
