# SetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.Chinese\_2010.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByValItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

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

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function sets the value of a concrete design overwrite item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignOverwriteItemChinese_2010()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



   'createSap2000 object

      Set SapObject= New Sap2000v16.SapObject



   'startSap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel= SapObject.SapModel



   'initialize model

      ret= SapModel.InitializeNewModel



   'create new concrete frame section property

      ret= SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret= SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret= SapModel.DesignConcrete.SetCode("Chinese 2010")



   'set overwrite item

      ret= SapModel.DesignConcrete.Chinese_2010.SetOverwrite("8", 1, 2)



   'closeSap2000

      SapObject.ApplicationExit False

      Set SapModel= Nothing

      Set SapObject= Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.2.

## See Also

[GetOverwrite](GetOverwrite_%7BConcrete_Chinese_2010%7D.md)
