# SetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.ACI318\_08\_IBC2009.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

Item

This is an integer between 1 and 15, inclusive, indicating the overwrite item considered.

1 = Framing type (beam and column)

2 = Live load reduction factor (beam and column)

3 = Unbraced length ratio, Major (beam and column)

4 = Unbraced length ratio, Minor (beam and column)

5 = Effective length factor, K Major (column only)

6 = Effective length factor, K Minor (column only)

7 = Moment coefficient, Cm Major (column only)

8 = Moment coefficient, Cm Minor (column only)

9 = Non-sway moment factor, Dns Major (column only)

10 = Non-sway moment factor, Dns Minor (column only)

11 = Sway moment factor, Ds Major (column only)

12 = Sway moment factor, Ds Minor (column only)

13 = Tangent of the angle of concrete compressive strut (beam only)

14 = Consider torsion (beam only)

15 = Concrete cover for closed stirrups (beam only)

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = Sway special

2 = Sway Intermediate

3 = Sway Ordinary

4 = Non-sway

2 = Live load reduction factor

Value >= 0; 0 means use program determined value.

3 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

4 = Unbraced length ratio, Minor

Value >= 0; 0 means use program determined value.

5 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

6 = Effective length factor, K Minor

Value >= 0; 0 means use program determined value.

7 = Moment coefficient, Cm Major

Value >= 0; 0 means use program determined value.

8 = Moment coefficient, Cm Minor

Value >= 0; 0 means use program determined value.

9 = Non-sway moment factor, Dns Major

Value >= 0; 0 means use program determined value.

10 = Non-sway moment factor, Dns Minor

Value >= 0; 0 means use program determined value.

11 = Sway moment factor, Ds Major

Value >= 0; 0 means use program determined value.

12 = Sway moment factor, Ds Minor

Value >= 0; 0 means use program determined value.

13 = Tangent of the angle of concrete compressive strut

Value >= 0; 0 means use program determined value.

14 = Consider torsion

0 = No

Any other value = Yes

15 = Concrete cover for closed stirrups

Value >= 0; 0 means use program determined value.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects=  2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function sets the value of a concrete design overwrite item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignOverwriteItemACI318_08_IBC2009()

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



       'create new concrete frame section property

          ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



       'set concrete design code

          ret = SapModel.DesignConcrete.SetCode("ACI318-08/IBC2009")



       'set overwrite item

          ret = SapModel.DesignConcrete.ACI318_08_IBC2009.SetOverwrite("8", 1, 4)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 14.2.0.

Added items 13~15 in version 23.4.0

## See Also

[GetOverwrite](mk:@MSITStore:W:/Released/SAP2000 V15/15.0.0 Server/CSi_OAPI_Documentation.chm::/SAP2000_API_Fuctions/Design/Concrete/ACI_318-05_IBC_2003/GetOverwrite_%7bConcrete_ACI_318-05_IBC_2003%7d.htm)
