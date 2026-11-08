# GetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.Italian\_DM\_14\_2\_92.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRefProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete frame design procedure.

Item

This is an integer between 1 and 12, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Live load reduction factor

3 = Unbraced length ratio, Major

4 = Unbraced length ratio, Minor

5 = Effective length factor, K Major

6 = Effective length factor, K Minor

7 = Amplification Coefficient, Omega Major

8 = Amplification Coefficient, Omega Minor

9 = Moment coefficient, c Major

10 = Moment coefficient, c Minor

11 = Moment coefficient, c\_sway Major

12 = Moment coefficient, c\_sway Minor

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = Sway

2 = Braced

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

7 = Amplification coefficient, Omega Major

Value >= 0; 0 means use program determined value.

8 = Amplification coefficient, Omega Minor

Value >= 0; 0 means use program determined value.

9 = Moment coefficient, c Major

Value >= 0; 0 means use program determined value.

10 = Moment coefficient, c Minor

Value >= 0; 0 means use program determined value.

11 = Moment coefficient, c\_sway Major

Value >= 0; 0 means use program determined value.

12 = Moment coefficient, c\_sway Minor

Value >= 0; 0 means use program determined value.

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a concrete design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignOverwriteItemItalian_DM_14_2_92()

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

      ret = SapModel.DesignConcrete.SetCode("Italian DM 14-2-92")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'get overwrite item

      ret = SapModel.DesignConcrete.Italian_DM_14_2_92.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetOverwrite](SetOverwrite_%7BConcrete_Italian_DM_14_292%7D.md)
