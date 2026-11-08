# GetTypeOAPI\_1

## Syntax

SapObject.SapModel.LoadCases.GetTypeOAPI\_1

## VB6 Procedure

Function GetTypeOAPI\_1(ByVal Name As String, ByRef CaseType As eLoadCaseType, ByRef SubType As Long, ByRef DesignType As eLoadPatternType, ByRef DesignTypeOption As Long, ByRef Auto As Long) As Long

## Parameters

Name

The name of an existing load case.

CaseType

This is one of the following items in the eLoadCaseType enumeration.

LinearStatic = 1

NonlinearStatic = 2

Modal = 3

ResponseSpectrum = 4

LinearHistory = 5  (Modal Time History)

NonlinearHistory = 6  (Modal Time History)

LinearDynamic = 7  (Direct Integration Time History)

NonlinearDynamic = 8  (Direct Integration Time History)

MovingLoad = 9

Buckling = 10

SteadyState = 11

PowerSpectralDensity = 12

LinearStaticMultistep = 13

Hyperstatic = 14

ExternalResults = 15

SubType

This is an integer representing the load case sub type. This item applies only for certain case types.

For CASE\_NONLINEAR\_STATIC:

1 = Nonlinear

2 = Nonlinear staged construction

3 = Nonlinear static multi-step

For CASE\_MODAL:

1 = Eigen

2 = Ritz

For CASE\_LINEAR\_HISTORY:

1 = Transient

2 = Periodic

DesignType

This is one of the following items in the eLoadPatternType enumeration.

Dead = 1

SuperDead = 2

Live = 3

ReduceLive = 4

Quake = 5

Wind = 6

Snow = 7

Other = 8

Move = 9

Temperature = 10

RoofLive = 11

Notional = 12

PatternLive = 13

Wave= 14

Braking = 15

Centrifugal = 16

Friction = 17

Ice = 18

WindOnLiveLoad = 19

HorizontalEarthPressure = 20

VerticalEarthPressure = 21

EarthSurcharge = 22

DownDrag = 23

VehicleCollision = 24

VesselCollision = 25

TemperatureGradient = 26

Settlement = 27

Shrinkage = 28

Creep = 29

WaterLoadPressure = 30

LiveLoadSurcharge = 31

LockedInForces = 32

PedestrianLL = 33

Prestress = 34

Hyperstatic = 35

Bouyancy = 36

StreamFlow = 37

Impact = 38

Construction = 39

DeadWearing = 40

DeadWater = 41

DeadManufacture = 42

EarthHydrostatic = 43

PassiveEarthPressure = 44

ActiveEarthPressure = 45

PedestrianLLReduced = 46

SnowHighAltitude = 47

EuroLm1Char = 48

EuroLm1Freq = 49

EuroLm2 = 50

EuroLm3 = 51

EuroLm4 = 52

DesignTypeOption

This is one of the following options for the DesignType item.

0 = Program determined

1 = User specified

Auto

This is one of the following values indicating if the load case has been automatically created.

0 = Not automatically created

1 = Automatically created by the bridge construction scheduler

## Remarks

This function retrieves the case type, design type, and auto flag for the specified load case.

The function returns zero if the type is successfully retrieved; otherwise it returns nonzero.

This function supersedes .GetType (LoadCases).

## VBA Example



```
Sub GetLoadCaseType_1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim CaseType As eLoadCaseType

      Dim SubType As Long

      Dim DesignType As Long

      Dim DesignTypeOption As Long

      Dim Auto As Long



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



   'get load case type

      ret = SapModel.LoadCases.GetTypeOAPI_1("DEAD", CaseType, SubType, DesignType, DesignTypeOption, Auto)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

This function supersedes [GetTypeOAPI {Load Case}](../../Obsolete_Functions/GetType_%7BLoad_Case%7D.md).

Changed function name to GetTypeOAPI in v17.0.0.

This function is obsolete and has been superseded by GetTypeOAPI\_2 as of v21.0.0. This function is maintained for backwards compatibility where staged construction and nonlinear multi-step cases are a subtype of load case type nonlinear static

## See Also
