# SetNZS11702004\_2

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetNZS11702004\_2

## VB6 Procedure

Function SetNZS11702004\_2(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal PeriodFlag As Long, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal NZS2004SiteClass As Long, ByVal NZS2004Z As Double, ByVal NZS2004R As Double, ByVal NZS2004DIST As Double, ByVal NZS2004Sp As Double, ByVal NZS2004Mu As Double, ByVal NZS2004ConsiderTSite As Boolean, ByVal NZS2004TSite As Double, ByVal NZS2004ConsiderSingleStory As Boolean) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

DirFlag

This is 1 or 2, indicating the seismic load direction.

1 = Global X

2 = Global Y

Eccen

The eccentricity ratio that applies to all diaphragms.

PeriodFlag

This is 1, 2 or 3, indicating the time period option.

1 = Approximate

2 = Program calculated

3 = User defined

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

**UserZ**

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

NZS2004SiteClass

This is 1, 2, 3, 4 or 5, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

NZS2004Z

The hazard factor, Z.

NZS2004R

The return period factor, R.

NZS2004DIST

Distance to the fault in km, used to calculate the near fault factor.

NZS2004Sp

The structural performance factor, Sp.

NZS2004Mu

The structural ductility factor, u.

NZS2004ConsiderTSite

Indicates whether to consider the site period for the spectral shape factor.

NZS2004TSite

The low amplitude site period.

NZS2004ConsiderSingleStory

Indicates whether to consider the structure as a single story, in which case Ft = 0.

## Remarks

This function assigns auto seismic loading parameters for the NZS 1170.5 2004 code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicNZS11702004_2()

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



       'create model from template

          ret = SapModel.File.New3DFrame(BeamSlab, 2, 144, 3, 336, 2, 432)



       'add new load pattern

          ret = SapModel.LoadPatterns.Add("EQX", LTYPE_QUAKE)



       'assign NZS 1170 2004 parameters

          ret = SapModel.LoadPatterns.AutoSeismic.SetNZS11702004_2("EQX", 2, 0.1, 2, 0, False, 0, 0, 3, 0.4, 1.3, 20, 0.7, 3, True, 1, False)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in v22.0.0.

This function supersedes SetNZS11702004\_1.

## See Also

GetNZS11702004\_2
