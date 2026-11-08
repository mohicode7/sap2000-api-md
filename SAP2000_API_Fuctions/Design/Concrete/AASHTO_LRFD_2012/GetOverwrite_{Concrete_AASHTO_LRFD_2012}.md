# GetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.AASHTO\_LRFD\_2012.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete frame design procedure.

Item

This is an integer between 1 and 11, inclusive, indicating the overwrite item considered.

1 = Live load reduction factor

2 = Unbraced length ratio, Major

3 = Unbraced length ratio, Minor

4 = Effective length factor, K Major

5 = Effective length factor, K Minor

6 = Moment coefficient, Cm Major

7 = Moment coefficient, Cm Minor

8 = Nonsway moment factor, Db Major

9 = Nonsway moment factor, Db Minor

10 = Sway moment factor, Ds Major

11 = Sway moment factor, Ds Minor

Value

The value of the considered overwrite item.

1 = Live load reduction factor

Value >= 0; 0 means use program determined value.

2 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

3 = Unbraced length ratio, Minor

Value >= 0; 0 means use program determined value.

4 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

5 = Effective length factor, K Minor

Value >= 0; 0 means use program determined value.

6 = Moment coefficient, Cm Major

Value >= 0; 0 means use program determined value.

7 = Moment coefficient, Cm Minor

Value >= 0; 0 means use program determined value.

8 = Nonsway moment factor, Db Major

Value >= 0; 0 means use program determined value.

9 = Nonsway moment factor, Db Minor

Value >= 0; 0 means use program determined value.

10 = Sway moment factor, Ds Major

  Value >= 0; 0 means use program determined value.

11 = Sway moment factor, Ds Minor

  Value >= 0; 0 means use program determined value.

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a concrete design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignOverwriteItemAASHTO_Concrete_LRFD_2012()

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

      ret = SapModel.InitializeNewModel



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("AASHTO LRFD 2012")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'get overwrite item

      ret = SapModel.DesignConcrete.AASHTO_LRFD_2012.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.3.0.

## See Also

[SetOverwrite](SetOverwrite_%7BConcrete_AASHTO_LRFD_2012%7D.md)
