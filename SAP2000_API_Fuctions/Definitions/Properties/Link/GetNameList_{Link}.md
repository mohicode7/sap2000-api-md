# GetNameList

## Syntax

SapObject.SapModel.PropLink.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String, Optional ByVal PropType As eLinkPropType) As Long

## Parameters

NumberNames

The number of link property names retrieved by the program.

MyName

This is a one-dimensional array of link property names. The MyName array is created as a dynamic, zero-based, array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the SAP2000 program, filled with the names, and returned to the API user.

PropType

This optional value is one of the following items in the eLinkPropType enumeration.

NLPROP\_LINEAR = 1

NLPROP\_DAMPER = 2

NLPROP\_GAP = 3

NLPROP\_HOOK = 4

NLPROP\_PLASTIC\_WEN = 5

NLPROP\_ISOLATOR1 = 6  (Rubber isolator)

NLPROP\_ISOLATOR2 = 7  (Friction isolator)

NLPROP\_MULTILINEAR\_ELASTIC = 8

NLPROP\_MULTILINEAR\_PLASTIC = 9

NLPROP\_ISOLATOR3 = 10  (T/C Friction isolator)

If no value is input for PropType, names are returned for all link properties in the model regardless of type.

## Remarks

This function retrieves the names of all defined link properties of the specified type.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetLinkPropNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

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



   'add link properties

      ReDim DOF(5)

      ReDim Fixed(5)

      ReDim Ke(5)

      ReDim Ce(5)

      DOF(0) = True

      Ke(0) = 12

      ret = SapModel.PropLink.SetLinear("L1", DOF, Fixed, Ke, Ce, 0, 0)

      ret = SapModel.PropLink.SetLinear("L2", DOF, Fixed, Ke, Ce, 0, 0)



   'get link property names

      ret = SapModel.PropLink.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
