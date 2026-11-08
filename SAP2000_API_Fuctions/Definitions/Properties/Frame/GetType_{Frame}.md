# GetTypeOAPI

## Syntax

SapObject.SapModel.PropFrame.GetTypeOAPI

## VB6 Procedure

Function GetTypeOAPI(ByVal Name As String, ByRef PropType
As eFramePropType) As Long

## Parameters

Name

The name of an existing frame section property.

PropType

This is one of the following items in the eFramePropType
enumeration.

I = 1

Channel = 2

T = 3

Angle = 4

DblAngle = 5

Box = 6

Pipe = 7

Rectangular = 8

Circle = 9

General = 10

DbChannel = 11

Auto = 12

SD = 13

Variable = 14

Joist = 15

Bridge = 16

Cold\_C = 17

Cold\_2C = 18

Cold\_Z = 19

Cold\_L = 20

Cold\_2L = 21

Cold\_HAT = 22

BuiltupICoverplate = 23

PCCGIRDERI = 24

PCCGIRDERU = 25

BuiltupIHybrid = 26

BuiltupUHybrid = 27

PCCGirderSuperT = 41

Cold\_Box = 42

Cold\_I = 43

Cold\_Pipe = 44

Cold\_T = 45

Trapezoidal = 46

## Remarks

This function retrieves the property type for the specified
frame section property.

The function returns zero if the type is successfully
retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetFramePropType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropType As eFramePropType



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'get frame property type

      ret = SapModel.PropFrame.GetTypeOAPI("FSEC1", 
 PropType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed function name to GetTypeOAPI in v17.0.0.

Added New PropType items in 23.1.0

## See Also
