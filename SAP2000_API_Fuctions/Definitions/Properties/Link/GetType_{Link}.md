# GetTypeOAPI

## Syntax

SapObject.SapModel.PropLink.GetTypeOAPI

## VB6 Procedure

Function GetTypeOAPI(ByVal Name As String, ByRef PropType As eLinkPropType) As Long

## Parameters

Name

The name of an existing link property.

PropType

This is one of the following items in the eLinkPropType enumeration.

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

## Remarks

This function retrieves the property type for the specified link property.

The function returns zero if the type is successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetLinkPropType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

      Dim PropType As eLinkPropType



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



   'add link property

      ReDim DOF(5)

      ReDim Fixed(5)

      ReDim Ke(5)

      ReDim Ce(5)

      DOF(0) = True

      Ke(0) = 12

      ret = SapModel.PropLink.SetLinear("L1", DOF, Fixed, Ke, Ce, 0, 0)



   'get link property type

      ret = SapModel.PropLink.GetTypeOAPI("L1", PropType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed function name to GetTypeOAPI in v17.0.0.

## See Also
