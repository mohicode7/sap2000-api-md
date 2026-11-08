# GetASCE716 (Note:  Obsolete, Newer Function Available)

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetASCE716

## VB6 Procedure

Function GetASCE716(ByVal Name As String, ByRef nDir() As Boolean, ByRef Eccen As Double, ByRef PeriodFlag As Integer, ByRef CtType As Integer, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef R As Double, ByRef Omega As Double, ByRef Cd As Double, ByRef I As Double, ByRef Ss As Double, ByRef S1 As Double, ByRef TL As Double, ByRef SiteClass As Integer, ByRef Fa As Double, ByRef Fv As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a ASCE 7-16 auto seismic load assignment.

nDir

This is an array with 2 inputs that indicate the seismic load direction.

nDir(1) = True = Global X

nDir(2) = True = Global Y

Eccen

The eccentricity ratio that applies to all diaphragms.

PeriodFlag

This is 1, 2 or 3, indicating the time period option.

1 = Approximate

2 = Program calculated

3 = User defined

CtType

This is 0, 1, 2 or 3, indicating the values of Ct and x. This item is meaningful when the PeriodFlag item is 1 or 2.

0 = Ct = 0.028 (ft),     x = 0.8

1 = Ct = 0.016 (ft),     x = 0.9

2 = Ct = 0.03 (ft),      x = 0.75

3 = Ct = 0.02 (ft),      x = 0.75

UserT

The user specified time period. This item is meaningful when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

R

The response modification factor.

Omega

The system overstrength factor.

Cd

The deflection amplification factor.

I

The occupancy importance factor.

SS, S1

The seismic coefficients Ss and S1.

TL

The long-period transition period. [s]

SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

Fa, Fv

The site coefficients Fa and Fv.

## Remarks

This function retrieves auto seismic loading parameters for the 2016 ASCE 7 code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersASCE716()



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create model from template

          ret = SapModel.File.New3DFrame(e3DFrameType.BeamSlab, 2, 144, 3, 336, 2, 432 )



       'add new load pattern

          ret = SapModel.LoadPatterns.Add("EQX", eLoadPatternType.Quake)
  
 
   'dimension ASCE716 parameters
        Dim nDir() As Boolean
        Dim Eccen As Double
        Dim PeriodFlag As Long
        Dim CtType As Long
        Dim UserT As Double
        Dim UserZ As Boolean
        Dim TopZ As Double
        Dim BottomZ As Double
        Dim R As Double
        Dim Omega As Double
        Dim Cd As Double
        Dim I As Double
        Dim SS As Double
        Dim S1 As Double
        Dim TL As Double
        Dim SiteClass As Long
        Dim Fa As Double
        Dim Fv As Double
 
        ReDim nDir(2)
        nDir(1) = True
        TopZ = 32
        BottomZ = 14
       'set ASCE716 parameters

          ret = SapModel.LoadPatterns.AutoSeismic.SetASCE716("EQX", nDir, 0.04, 3, 1, 1.76, True, TopZ, BottomZ, 6, 3.5, 6.5, 1.5, 1.9, 1.1, 8, 3, 0, 0)



       'get ASCE716 parameters

          ret = SapModel.LoadPatterns.AutoSeismic.GetASCE716("EQX", nDir, Eccen, PeriodFlag, CtType, UserT, UserZ, TopZ, BottomZ, R, Omega, Cd, I, SS, S1, TL, SiteClass, Fa, Fv)



          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 24.2.0.

This function was replaced by [GetASCE716\_1](../Definitions/Load_Pattern/Auto_Seismic_Load/GetASCE_16_%7BAuto_Seismic%7D.md), which is more consistent with other API functions.

## See Also

[SetASCE716](SetASCE16_%7BAuto_Seismic%7D-obsolete.md)
