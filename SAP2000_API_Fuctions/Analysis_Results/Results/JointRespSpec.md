# JointRespSpec

## Syntax

SapObject.SapModel.Results.JointRespSpec

## VB6 Procedure

Function JointRespSpec(ByVal Name As String, ByVal ItemTypeElm As eItemTypeElm, ByVal edSet As String, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef LoadCase() As String, ByRef CoordSys() As String, ByRef Dir() As Long, ByRef Damping() As Double, ByRef SpecWidening() As Double, ByRef AbscissaValue() As Double, ByRef OrdinateValue() As Double) As Long

## Parameters

Name

The name of an existing point object, point element, or group of objects depending on the value of the ItemTypeElm item.

ItemTypeElem

This is one of the following items in the ItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

Default = 4

If this item is ObjectElm, the result request is for the point element corresponding to the point object specified by the Name item.

If this item is Element, the result request is for the point element specified by the Name item.

If this item is GroupElm, the result request is for all point elements directly or indirectly specified in the group specified by the Name item.

If this item is SelectionElm, the result request is for all point elements directly or indirectly selected and the Name item is ignored.

If this item is Default, the result request is for the point object specified in the named set.

See [Item Type for Elements](Item_Type_for_Elements.md) for more information.

NamedSet

The name of an existing joint response spectrum named set. See SapModel.NamedSet.SetJointRespSpec.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the point object name associated with each result, if any. Some results will have no point object associated with them. For those cases, this item will be blank.

Elm

This is an array that includes the point element name associated with each result.

LoadCase

This is an array that includes the name of the load case associated with each result.

CoordSys

This is an array that includes the coordinate system in which the results are reported.

Dir

This is an array that includes the direction for which the results are reported. Valid values for the direction are:

1 = Local 1, Global X, or user-defined coordinate system X

2 = Local 2, Global Y, or user-defined coordinate system Y

3 = Local 3, Global Z, or user-defined coordinate system Z

Damping

This is an array that includes the critical damping ratio, for each result.

SpecWidening

This is an array that includes the percent spectrum widening, for each result.

AbscissaValue

This is an array that includes the period or frequency, as defined in each named set, for each result. [s or 1/s]

OrdinateValue

This is an array of the response quantity, as defined in each named set, for each result. The possible response quantities are spectral displacement [L], spectral velocity [L/s], pseudo spectral velocity [L/s], spectral acceleration [L/s2], or pseudo spectral acceleration [L/s2].

## Remarks

This function reports the joint response spectra values, due to a time history analysis, for the specified point elements.

The function returns zero if the response spectra data is successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information

## VBA Example



```
Sub GetJointResponseSpectra()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim MyLoadType() As String 
      Dim MyLoadName() As String 
      Dim MyFunc() As String 
      Dim MySF() As Double 
      Dim MyTF() As Double 
      Dim MyAT() As Double 
      Dim MyCSys() As String 
      Dim MyAng() As Double 
      Dim Joints() As String 
      Dim UserFreq() As Double 
      Dim DampValues() As Double 
      Dim NumberResults As Long 
      Dim Obj() As String 
      Dim Elm() As String 
      Dim LoadCase() As String 
      Dim CoordSys() As String 
      Dim Dir() As Long 
      Dim Damping() As Double 
      Dim SpecWidening() As Double 
      Dim AbscissaValue() As Double 
      Dim OrdinateValue() As Double



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
 
   'add sine TH function

      ret = SapModel.Func.FuncTH.SetSine("TH-1, 1, 16, 4, 1.25)



   'add linear modal history load case

      ReDim MyLoadType(0)

      ReDim MyLoadName(0)

      ReDim MyFunc(0)

      ReDim MySF(0)

      ReDim MyTF(0)

      ReDim MyAT(0)

      ReDim MyCSys(0)

      ReDim MyAng(0)

      MyLoadType(0) = "Accel"

      MyLoadName(0) = "U1"

      MyFunc(0) = "TH-1"

      MySF(0) = 2

      MyTF(0) = 1.5

      MyAT(0) = 10

      MyCSys(0) = "Global"

      MyAng(0) = 10
 
      ret = SapModel.LoadCases.ModHistLinear.SetCase("LCASE1")
      ret = SapModel.LoadCases.ModHistLinear.SetLoads("LCASE1", 1, MyLoadType, MyLoadName, MyFunc, MySF, MyTF, MyAT, MyCSys, MyAng)
 
   'define named set

      ReDim Joints(1)

      Joints(0) = "3"

      Joints(1) = "6"

      ReDim DampValues(1)

      DampValues(0) = 0

      DampValues(1) = 0.05



      ret = SapModel.NamedSets.SetJointRespSpec("Sample", "LCASE1", 2, Joints, "Global", 1, 1, 4, True, True, 0, UserFreq, 2, DampValues)
 
   'run analysis
      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis
 
   'clear all case and combo output selections and select time history case

      ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



   'set case and combo output selections

      ret = SapModel.Results.Setup.SetCaseSelectedForOutput("LCASE1")
   'get point response spectra
      ret = SapModel.Results.JointRespSpec("3", 4, "Sample", NumberResults, Obj, Elm, LoadCase, CoordSys, Dir, Damping, SpecWidening, AbscissaValue, OrdinateValue)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.3.0.

## See Also

[SetJointRespSpec](../../Definitions/Named_Sets/SetJointRespSpec.md)

[JointDisplAbs](JointDisplAbs.md)

[JointVel](JointVel.md)

[JointVelAbs](JointVelAbs.md)

[JointAcc](JointAcc.md)

[JointAccAbs](JointAccAbs.md)

[GeneralizedDispl](GeneralizedDispl.md)
