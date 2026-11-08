# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.Australian\_AS4100\_2020.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value
As Double) As Long

## Parameters

Item

This is an integer between 1 and 17, inclusive, indicating
the preference item considered.

1 = Multi-response case
design

2 = Framing type

3 = Structural analysis
method

4 = Steel type

5 = Capacity factor, Phi
bending

6 = Capacity factor, Phi
compression

7 = Capacity factor, Phi
tension yielding

8 = Capacity factor, Phi
tension fracture

9 = Capacity factor, Phi
shear

10 = Consider deflection

11 = DL deflection limit,
L/Value

12 = SDL + LL deflection
limit, L/Value

13 = LL deflection limit,
L/Value

14 = Total load deflection
limit, L/Value

15 = Total camber limit,
L/Value

16 = Pattern live load
factor

17 = Demand/capacity ratio
limit

Value

The value of the considered preference item.

1 = Multi-response case
design

1 = Envelopes

2 = Step-by step

3 = Last step

4 = Envelopes - All

5 = Step-by step - All

2 = Framing type

1 = Moment frame

2 = Braced frame

3 = Structural analysis
method

1 = General 2nd Order

2 = Amplified 1st Order

4 = Steel type

1 = Hot rolled

2 = Hot finished

3 = Cold form

4 = Stress relieved

5 = Lightly welded

6 = Heavily welded

5 = Capacity factor, Phi
bending

Value > 0

6 = Capacity factor, Phi
compression

Value > 0

7 = Capacity factor, Phi
tension yielding

Value > 0

8 = Capacity factor, Phi
tension fracture

Value > 0

9 = Capacity factor, Phi
shear

Value > 0

10 = Consider deflection

  0 = No

  Any other
value = Yes

11 = DL deflection limit,
L/Value

  Value >
0

12 = SDL + LL deflection
limit, L/Value

  Value >
0

13 = LL deflection limit,
L/Value

  Value >
0

14 = Total load deflection
limit, L/Value

  Value >
0

15 = Total camber limit,
L/Value

  Value >
0

16 = Pattern live load
factor

  Value >=
0

17 = Demand/capacity ratio
limit

  Value >
0

## Remarks

This function retrieves the value of a steel design
preference item.

The function returns zero if the item is successfully
retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItem Australian_AS4100_2020()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("AS 
 4100-2020")



   'get preference item

      ret = SapModel.DesignSteel.Australian_AS4100_2020.GetPreference(1, 
 Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 25.0

## See Also

[SetPreference](SetPreference_%7BSteel_AS_4100-2020%7D.md)
