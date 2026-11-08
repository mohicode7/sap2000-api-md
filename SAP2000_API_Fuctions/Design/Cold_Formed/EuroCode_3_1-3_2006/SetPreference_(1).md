# **SetPreference**

## **Syntax**

SapObject.SapModel.Design<a id="ColdFormed"></a><a id="ColdFormed"></a>ColdFormed.EuroCold06.SetPreference

## **VB6 Procedure**

Function SetPreference(ByVal Item As Long, ByVal Value
As Double) As Long

## **Parameters**

Item

This is an integer between 1 and 12, inclusive, indicating
the preference item considered.

1 = Country

2 = Multi-response case design

3 = Demand/capacity ratio limit

4 = Combos equation

5 = Reliability class

6 = K factor method

7 = Consider P-Delta Done

8 = GammaM0

9 = GammaM1

10 = GammaM2

11 = Pattern live load factor

12 = Member buckling interaction equations

Value

The value of the considered preference
item.

1 = Country

     1 = CEN
Default

     2 = United
Kingdom

     3 = Slovenia

     4 = Bulgaria

     5 = Norway

     7 = Sweden

     8 = Finland

     9 = Denmark

   10 = Portugal

   11 = Germany

2 = Multi-response case design

1 = Envelopes

2 = Step-by-step

3 = Last step

4 = Envelopes -- All

5 = Step-by-step --  All

3 = Demand/capacity ratio limit

Value > 0

4 = Combos equation

      1 = Eq.
6.10

      2 = Max
of Eqs. 6.10a and 6.10b

5 = Reliability class

     1 = Class 1

     2 = Class 2

     3 = Class 3

6 = K factor method

     1 = Method 1 (Annex
A)

     2 = Method 2 (Annex
B)

   

7 = Consider P-Delta Done

0 = No

Any other value = Yes

  

8 = GammaM0

Value > 0

9 = GammaM1

Value > 0

10 = GammaM2

       Value
> 0

  

11 = Pattern live load factor

 Value >= 0

12 = Member buckling interaction
equations

1 = Equations 6.61 and
6.62 in EN 1993-1-1:2005

2 = Equation 6.36 in EN
1993-1-3:2006 for each axis of bending separately

3
= Equation 6.36 in EN 1993-1-3:2006 for both axes of bending simultaneously

4
= Equation 6.61 and 6.62 in EN 1993-1-1:2005 and Equation 6.36 in EN 1993-1-3:2006
for each axis of bending separately

5
= Equation 6.61 and 6.62 in EN 1993-1-1:2005 and Equation 6.36 in EN 1993-1-3:2006
for both axes of bending simultaneously

## **Remarks**

This function sets the value of a cold-formed steel
design preference item.

The function returns zero if the item is successfully
set; otherwise, it returns a nonzero value.

## **VBA Example**



```
Sub SetSteelDesignPreferenceItemEuroCold06()

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

      ret = SapModel.DesignColdFormed.SetCode("Eurocode 
 3 1-3 2006")



   'set preference item

      ret = SapModel.DesignColdFormed.EuroCold06.SetPreference(4, 
 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## **Release Notes**

Initial release in version 23.0.0.

Added item 12 in version 23.4.0

## **See Also**

[GetPreference](GetPreference_%281%29.md)
