# SetPreference

## Syntax

SapObject.SapModel.DesignConcrete.Chinese\_2002.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 11, inclusive, indicating the preference item considered.

1 = Number of interaction curves

2 = Number of interaction points

3 = Importance factor gamma 0

4 = Column design procedure

5 = Seismic design grade

6 = Pattern live load factor

7 = Utilization factor limit

8 = Multi-response case design

9 = Structural system

10 = Is tall building?

11 = Seismic field type

Value

The value of the considered preference item.

1 = Number of interaction curves

Value >= 4 and devisable by 4

2 = Number of interaction points

Value >= 5 and odd

3 = Importance factor gamma 0

Value > 0

4 = Column design procedure

1 = Appendix F

2 = Simplified

5 = Seismic design grade

1 = Super I

2 = Grade I

3 = Grade II

4 = Grade III

5 = Grade IV

6 = Nonseismic

6 = Pattern live load factor

Value >= 0

7 = Utilization factor limit

Value > 0

8 = Multi-response case design

1 = Envelopes

2 = Step-by-step

3 = Last step

4 = Envelopes -- All

5 = Step-by-step -- All

9 = Structural system

      1 = Frame only

      2 = Shearwall only

      3 = Frame-shearwall

      4 = Braced frame only

      5 = Frame-braced frame

10 =  Is tall building?

        0 = No

        1 = Yes

11 =  Seismic field type

        1 = I

        2 = II

        3 = III

        4 = IV

## Remarks

This function sets the value of a concrete design preference item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignPreferenceItemChinese_2002()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



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

      ret = SapModel.DesignConcrete.SetCode("Chinese 2002")



   'set preference item

      ret = SapModel.DesignConcrete.Chinese_2002.SetPreference(2, 9)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Added Items 9, 10, and 11 in Version 14.0.0.

Changed Time history design item to Multi-response case design and added additional values in version 15.0.1.

## See Also

[GetPreference](GetPreference_%7BConcrete_Chinese_2002%7D.md)
