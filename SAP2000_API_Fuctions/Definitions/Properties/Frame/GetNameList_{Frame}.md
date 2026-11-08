# GetNameList

## Syntax

SapObject.SapModel.PropFrame.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String, Optional ByVal PropType As eFramePropType) As Long

## Parameters

NumberNames

The number of frame section property names retrieved by the program.

MyName

This is a one-dimensional array of frame section property names. The MyName array is created as a dynamic, zero-based, array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the SAP2000 program, filled with the names, and returned to the API user.

PropType

This optional value is one of the following items in the eFramePropType enumeration.

SECTION\_I = 1

SECTION\_CHANNEL = 2

SECTION\_T = 3

SECTION\_ANGLE = 4

SECTION\_DBLANGLE = 5

SECTION\_BOX = 6

SECTION\_PIPE = 7

SECTION\_RECTANGULAR = 8

SECTION\_CIRCLE = 9

SECTION\_GENERAL = 10

SECTION\_DBCHANNEL = 11

SECTION\_AUTO = 12

SECTION\_SD = 13

SECTION\_VARIABLE = 14

SECTION\_JOIST = 15

SECTION\_BRIDGE = 16

SECTION\_COLD\_C = 17

SECTION\_COLD\_2C = 18

SECTION\_COLD\_Z = 19

SECTION\_COLD\_L = 20

SECTION\_COLD\_2L = 21

SECTION\_COLD\_HAT = 22

SECTION\_BUILTUP\_I\_COVERPLATE = 23

SECTION\_PCC\_GIRDER\_I = 24

SECTION\_PCC\_GIRDER\_U = 25  
SECTION\_BUILTUP\_I\_HYBRID = 26

SECTION\_BUILTUP\_U\_HYBRID = 27

If no value is input for PropType, names are returned for all frame section properties in the model regardless of type.

## Remarks

This function retrieves the names of all defined frame section properties of the specified type.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetFramePropNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberNames As Long

      Dim MyName() As String



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



   'get frame section property names

      ret = SapModel.PropFrame.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Builtup I Hybrid and Builitup U Hybrid added with Version 15.0.0.

## See Also
