# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.Canadian\_S16\_24.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 21, inclusive, indicating the preference item considered.

1 = Multi-response case design

2 = Framing type

3 = Spectral Acceleration Ratio, Ie\*Fa\*Sa(0.2)

4 = Ductility related modification factor, Rd

5 = Overstrength related modification factor, Ro

6 = Capacity factor, Phi bending

7 = Capacity factor, Phi compression

8 = Capacity factor, Phi tension

9 = Capacity factor, Phi shear

10 = Ignore seismic code

11 = Ignore special seismic load

12 = Doubler plate is plug welded

13 = Consider deflection

14 = DL deflection limit, L/Value

15 = SDL + LL deflection limit, L/Value

16 = LL deflection limit, L/Value

17 = Total load deflection limit, L/Value

18 = Total camber limit, L/Value

19 = Pattern live load factor

20 = Demand/capacity ratio limit

21 = Seismic Design Category

Value

The value of the considered preference item.

1 = Multi-response case design

1 = Envelopes

2 = Step-by step

3 = Last step

4 = Envelopes - All

5 = Step-by step - All

2 = Framing type

1 = Type LD MRF

2 = Type MD MRF

3 = Type D MRF

4 = Type LD CBF(V)

5 = Type LD CBF(TC)

6 = Type LD CBF(TO)

7 = Type LD CBF(OT)

8 = Type MD CBF(V)

9 = Type MD CBF(TC)

10 = Type MD CBF(TO)

11 = Type MD CBF(OT)

12 = EBF

13 = Cantilever Column

14 = Conventional MF

15 = Conventional BF

3 = Spectral Acceleration Ratio, Ie\*Fa\*Sa(0.2)

Value > 0

4 = Ductility related modification factor, Rd

Value > 0

5 = Overstrength related modification factor, Ro

Value > 0

6 = Capacity factor, Phi bending

Value > 0

7 = Capacity factor, Phi compression

Value > 0

8 = Capacity factor, Phi tension

Value > 0

9 = Capacity factor, Phi shear

Value > 0

 

10 = Ignore seismic code

  0 = No

  Any other value = Yes  
 

11 = Ignore special seismic load

  0 = No

  Any other value = Yes  
 

12 = Doubler plate is plug welded

  0 = No

  Any other value = Yes

13 = Consider deflection

  0 = No

  Any other value = Yes

14 = DL deflection limit, L/Value

  Value > 0

15 = SDL + LL deflection limit, L/Value

  Value > 0

16 = LL deflection limit, L/Value

  Value > 0

17 = Total load deflection limit, L/Value

  Value > 0

1 = Total camber limit, L/Value

  Value > 0

19 = Pattern live load factor

  Value >= 0

20 = Demand/capacity ratio limit

  Value > 0

21 = Seismic Design Category

  1 = SC0

  2 = SC1

  3 = SC2

  4 = SC3

  5 = SC4

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemCanadian_S16_24 ()

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

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



       'set steel design code

          ret = SapModel.DesignSteel.SetCode("CSA S16-24")



       'get preference item

          ret = SapModel.DesignSteel.Canadian_S16_24.GetPreference(1, Value)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 27.0.0

## See Also

[SetPreference](SetPreference_%7BSteel_CSA_S16-24%7D.md)
