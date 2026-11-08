# SetDesignType

## Syntax

SapObject.SapModel.LoadCases.SetDesignType

## VB6 Procedure

Function SetDesignType(ByVal Name As String, ByVal DesignTypeOption As Long, Optional ByVal DesignType As eLoadPatternType = LTYPE\_DEAD) As Long

## Parameters

Name

The name of an existing load case.

DesignTypeOption

This is one of the following options for the DesignType item.

0 = Program determined

1 = User specified

DesignType

This item only applies when the DesignTypeOption is 1 (user specified). It is one of the following items in the eLoadPatternType enumeration.

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

This function sets the design type for the specified load case.

The function returns zero if the type is successfully set; otherwise it returns nonzero.

## VBA Example



```
Sub SetLoadCaseDesignType



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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set load case design type

      ret = SapModel.LoadCases.SetDesignType("DEAD", 1, LTYPE_SUPERDEAD)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetType\_1](GetType_1_%7BLoad_Case%7D.md)
