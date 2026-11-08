# GetLoadType

## Syntax

SapObject.SapModel.LoadPatterns.GetLoadType

## VB6 Procedure

Function GetLoadType(ByVal Name As String, ByRef MyType As eLoadPatternType) As Long

## Parameters

Name

The name of an existing load pattern.

MyType

This is one of the following items in the eLoadPatternType enumeration:

LTYPE\_DEAD = 1

LTYPE\_SUPERDEAD = 2

LTYPE\_LIVE = 3

LTYPE\_REDUCELIVE = 4

LTYPE\_QUAKE = 5

LTYPE\_WIND= 6

LTYPE\_SNOW = 7

LTYPE\_OTHER = 8

LTYPE\_MOVE = 9

LTYPE\_TEMPERATURE = 10

LTYPE\_ROOFLIVE = 11

LTYPE\_NOTIONAL = 12

LTYPE\_PATTERNLIVE = 13

LTYPE\_WAVE= 14

LTYPE\_BRAKING = 15

LTYPE\_CENTRIFUGAL = 16

LTYPE\_FRICTION = 17

LTYPE\_ICE = 18

LTYPE\_WINDONLIVELOAD = 19

LTYPE\_HORIZONTALEARTHPRESSURE = 20

LTYPE\_VERTICALEARTHPRESSURE = 21

LTYPE\_EARTHSURCHARGE = 22

LTYPE\_DOWNDRAG = 23

LTYPE\_VEHICLECOLLISION = 24

LTYPE\_VESSELCOLLISION = 25

LTYPE\_TEMPERATUREGRADIENT = 26

LTYPE\_SETTLEMENT = 27

LTYPE\_SHRINKAGE = 28

LTYPE\_CREEP = 29

LTYPE\_WATERLOADPRESSURE = 30

LTYPE\_LIVELOADSURCHARGE = 31

LTYPE\_LOCKEDINFORCES = 32

LTYPE\_PEDESTRIANLL = 33

LTYPE\_PRESTRESS = 34

LTYPE\_HYPERSTATIC = 35

LTYPE\_BOUYANCY = 36

LTYPE\_STREAMFLOW = 37

LTYPE\_IMPACT = 38

LTYPE\_CONSTRUCTION = 39

## Remarks

This function retrieves the load type for a specified load pattern.

The function returns zero if the load type is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLoadPatternType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As eLoadPatternType



   'create Sap2000 object

      Set SapObject = New Sap2000v16.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'get load pattern type

      ret = SapModel.LoadPatterns.GetLoadType("DEAD", MyType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Added three items to the eLoadPatternType enumeration in version 12.00.

## See Also

[SetLoadType](SetLoadType.md)
