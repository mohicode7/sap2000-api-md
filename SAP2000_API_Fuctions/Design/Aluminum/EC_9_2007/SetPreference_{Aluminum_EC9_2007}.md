# SetPreference

## Syntax

SapObject.SapModel.DesignAluminum.Eurocode\_9\_2007.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value
As Double) As Long

## Parameters

Item

This is an integer between 1 and 10, inclusive, indicating
the preference item considered.

1
= Country

2 = Multi-response case
design

3 = Demand/capacity ratio
limit

4 = Combos equation

5 = Reliability class

6 = GammaM1

7 = GammaM2

8 = Consider P-Delta Done

9 = Consider Torsion

10 = Pattern live load
factor

Value

The value of the considered preference item.

1 = Country

      1
= CEN Default

      2
= United Kingdom

      3
= Slovenia

      4
= Bulgaria

      5
= Norway

      7
= Sweden

      8
= Finland

      9
= Denmark

     10
= Portugal

     11
= Germany

2 = Multi-response case
design

1 = Envelopes

2 = Step-by-step

3 = Last step

4 = Envelopes -- All

5 =
Step-by-step -- All

3 = Demand/capacity ratio
limit

  Value >
0

4 = Combos equation

   1 = Eq.
6.10

   2 = Max
of Eqs. 6.10a and 6.10b

5 = Reliability class

   1 = Class
1

   2 = Class
2

   3 = Class
3

       

6 = GammaM1

Value > 0

7 = GammaM2

      Value
> 0

8 = Consider P-Delta Done

0 = No

Any other value = Yes

9 = Consider Torsion

0 = No

Any other value = Yes

  

10 = Pattern live load
factor

  Value >=
0

## Remarks

This function sets the value of an aluminum design preference
item.

The function returns zero if the item is successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAluminumDesignPreferenceItemEurocode_9_2007()

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



   'set aluminum design code

      ret = SapModel.DesignAluminum.SetCode("EN 
 1999:2007")



   'set preference item

      ret = SapModel.DesignAluminum.Eurocode_9_2007.SetPreference(4, 
 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.3.0

## See Also

[GetPreference](GetPreference_%7BAluminum_EC9_2007%7D.md)
