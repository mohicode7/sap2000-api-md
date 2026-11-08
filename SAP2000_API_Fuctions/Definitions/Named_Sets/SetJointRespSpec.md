# SetJointRespSpec

## Syntax

SapObject.SapModel.NamedSet.SetJointRespSpec

## VB6 Procedure

Function SetJointRespSpec(ByVal Name As String, ByVal LoadCase As String, ByVal NumberJoints As Long, ByVal JointNames() As String, ByVal CoordSys As String, ByVal Direction As Long, ByVal Abscissa As Long, ByVal Ordinate As Long, ByVal DefaultFreq As Boolean, ByVal StructuralFreq As Boolean, ByVal NumberUserFreq As Long, ByVal UserFreqValues() As Double, ByVal NumberDampValues As Long, ByVal DampingValues() As Double, Optional ByVal AbscissaPlotType As Long = 2, Optional ByVal SpectrumWidening As Double = 0, Optional ByVal OrdinatePlotType As Long = 1, Optional ByVal OrdinateScaleFactor As Double = 1) As Long

## Parameters

Name

The name of a new or existing joint response spectrum named set. If the named set already exists, it will be modified, otherwise a new named set will be created.

Load Case

The name of a time history load case for which results will be extracted.

NumberJoints

The number of joints for which to generate response spectra.

JointNames

This is an array of joint names for which to generate response spectra.

CoordSys

The name of a coordinate system in which the direction for results is defined. This can be Global, Local, or any other user-defined coordinate system.

Direction

This is a value specifying the direction, in the specified coordinate system, in which the results are to be retrieved. Valid values for the direction are:

1 = Local 1, Global X, or user-defined coordinate system X  
2 = Local 2, Global Y, or user-defined coordinate system Y  
3 = Local 3, Global Z, or user-defined coordinate system Z

Abscissa

This is one of the following, specifying the abscissa data type.

1 = Frequency  
2 = Period

Ordinate

This is one of the following, specifying the ordinate data type.

1 = Spectral displacement  
2 = Spectral velocity  
3 = Pseudo-spectral velocity  
4 = Spectral acceleration  
5 = Pseudo-spectral acceleration

DefaultFreq

If this item is True, the default frequencies are used for output.

StructuralFreq

If this item is True, the structural frequencies are used for output.

NumberUserFreq

The number of user-defined frequencies, which may be 0.

UserFreqValues

This is an array that includes the user-defined frequencies. Values are returned  using indices 0 to NumberUserFreq-1 if NumberUserFreq > 0. [1/s]

NumberDampValues

The number of critical damping ratio values, which must be at least 1..

Damping Values

This is an array that includes the critical damping ratios to be used for extracting results. Values are supplied using indices 0 to NumberDampValues-1.

Abscissa Plot Type

This is one of the following, specifying the abscissa axis scale type.

1 = Arithmetic

2 = Log

Spectrum Widening

This specifies the percentage by which to widen the peaks of the spectrum.

OrdinatePlotType

This is one of the following, specifying the ordinate axis scale type.

1 = Arithmetic

2 = Log

OrdinateScaleFactor

This is the scale factor used to linearly scale the response spectrum ordinate values.

## Remarks

This function defines a joint response spectrum named set.

The function returns zero if the named set is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub DefineJointRSNamedSet()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim Joints() As String

      Dim UserFreq() As Double
      Dim DampValues() As Double

      

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
 
  'add linear modal history load case

      ret = SapModel.LoadCases.ModHistLinear.SetCase("LCASE1")



   'define named set

      ReDim InJoints(1)

      Joints(0) = "3"

      Joints(1) = "6"

      ReDim DampValues(1)

      DampValues(0) = 0

      DampValues(1) = 0.05



   ret = SapModel.NamedSet.SetJointRespSpec("Sample", "LCASE1", 2, Joints, "Global", 1, 1, 4, True, True, 0, UserFreq, 2, DampValues)
 
   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.3.0.

## See Also

[GetJointRespSpec](GetJointRespSpec.md)
