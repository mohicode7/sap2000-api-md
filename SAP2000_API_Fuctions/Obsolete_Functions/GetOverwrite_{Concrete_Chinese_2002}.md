# GetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.Chinese\_2002.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete frame design procedure.

Item

This is an integer between 1 and 18, inclusive, indicating the overwrite item considered.

1 = Seismic design grade

2 = Dual system SMF

3 = MMF

4 = SMF

5 = AFMF

6 = Column location

7 = Transfer beam of column

8 = Corner column seismic modification

9 = Beam gravity neg moment red factor

10 = Unbraced length ratio, Major

11 = Unbraced length ratio, Minor

12 = Effective length factor, K Major

13 = Effective length factor, K Minor

14 = Torsion modification factor

15 = Torsion design factor, Zeta

16 = Concrete cover for closed stirrup

17 = Effective length factor for gravity, K Major

18 = Effective length factor for gravity, K Minor

Value

The value of the considered overwrite item.

1 = Seismic design grade

0 = As specified in preferences

1 = Seismic Super I

2 = Seismic Class I

3 = Seismic Class II

4 = Seismic Class III

5 = Seismic Class IV

6 = NonSeismic

2 = Dual system SMF

Value >= 0; 0 means use program determined value.

3 = MMF

Value >= 0; 0 means use program determined value.

4 = SMF

Value >= 0; 0 means use program determined value.

5 = AFMF

Value >= 0; 0 means use program determined value.

6 = Column Location

1 = Center Column

2 = Side Column

3 = Corner Column

4 = End Column

5 = Individual Column

7 = Transfer beam or column

0 = Program Determined

1 = No

2 = Yes

8 = Corner column seismic modification

0 = Program Determined

1 = No

2 = Yes

9 = Beam gravity neg moment red factor

Value >= 0; 0 means use program determined value.

10 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

11 = Unbraced length ratio, Minor

Value >= 0; 0 means use program determined value.

12 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

13 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

14 = Torsion modification factor

Value >= 0; 0 means use program determined value.

15 = Torsion design factor, Zeta

Value >= 0; 0 means use program determined value.

16 = Concrete cover for closed stirrup

Value >= 0; 0 means use program determined value.

17 = Effective length factor for gravity, K Major

        Value >= 0; 0 means use program default value.

18 = Effective length factor for gravity, K Minor

        Value >= 0; 0 means use program default value.

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a concrete design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignOverwriteItemChinese_2002()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Double

      Dim ProgDet As Boolean



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



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'get overwrite item

      ret = SapModel.DesignConcrete.Chinese_2002.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Added Items 17 and 18 in Version 14.0.0.

## See Also

[SetOverwrite](SetOverwrite_%7BConcrete_Chinese_2002%7D.md)
