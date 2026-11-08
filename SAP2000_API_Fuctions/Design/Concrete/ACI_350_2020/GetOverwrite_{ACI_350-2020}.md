# GetOverwrite

## Syntax

SapObject.SapModel.DesignConcreteShell.ACI350\_20.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete shell design procedure.

Item

This is an integer between 1 and 11, inclusive, indicating the overwrite item considered.

1 = Number rebar layer

2 = Concrete cover to center of top rebar layer in direction 1

3 = Concrete cover to center of top rebar layer in direction 2

4 = Concrete cover to center of bottom rebar layer in direction 1

5 = Concrete cover to center of bottom rebar layer in direction 2

6 = Longitudinal bar size

7 = Environmental exposure conditions

8 = Environmental durability factor for axial forces

9 = Environmental durability factor for moments

10 = Environmental durability factor for shear

11 = Consider shear design

Value

The value of the considered overwrite item.

1 = Number rebar layer

0 = Program Default

1 = 1 layer

2 = 2 layers

2 = Concrete cover to center of top rebar layer in direction 1

Value >= 0; 0 means value taken from shell section definition.

3 = Concrete cover to center of top rebar layer in direction 2

Value >= 0; 0 means value taken from shell section definition.

4 = Concrete cover to center of bottom rebar layer in direction 1

Value >= 0; 0 means value taken from shell section definition.

5 = Concrete cover to center of bottom rebar layer in direction 2

Value >= 0; 0 value taken from shell section definition.

6 = Longitudinal bar size

      1 = #2

      2 = #3

      3 = #4

      4 = #5

      5 = #6

      6 = #7

      7 = #8

      8 = #9

      9 = #10

      10 = #11

      11 = #14

      12 = #18

      13 = 10M

      14 = 15M

      15 = 20M

      16 = 25M

      17 = 30M

      18 = 35M

      19 = 45M

      20 = 55M

      21 = 6d

      22 = 8d

      23 = 10d

      24 = 12d

      25 = 14d

      26 = 16d

      27 = 20d

      28 = 25d

      29 = 26d

      30 = 28d

      31 = N12

      32 = N16

      33 = N20

      34 = N24

      35 = N28

      36 = N32

      37 = N36

7 = Environmental exposure conditions

0 = Program Default

1 = Normal

2 = Extreme

8 = Environmental durability factor for axial forces

Value >= 0; 0 means use program determined value.

9 = Environmental durability factor for moments

Value >= 0; 0 means use program determined value

10 = Environmental durability factor for shear

Value >= 0; 0 means use program determined value

11 = Consider shear design

 0 = No

Any other value = Yes

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a concrete design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcretShellDesignOverwriteItemACI350_20()
   'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim Value As Double

          Dim ProgDet As Boolean



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model
  
      ret = SapModel.InitializeNewModel(eUnits.kip_in_F)

   'create a wall model from template
      ret = SapModel.File.NewWall(6, 4, 6, 4)
       'set concrete shell design code

          ret = SapModel.DesignConcreteShell.SetCode("ACI 350-20")

   'set concrete shell design overwrite
      ret = SapModel.DesignConcreteShell.ACI350_20.SetOverwrite("1", 4, 1.345)



       'get overwrite item

          ret = SapModel.DesignConcreteShell.ACI350_20.GetOverwrite("1", 4, value, progdet)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 26.0.0.

## See Also

[SetOverwrite](SetPreference_%7BACI_350-2020%7D.md)
