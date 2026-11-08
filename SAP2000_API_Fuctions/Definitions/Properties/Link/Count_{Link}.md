# Count

## Syntax

SapObject.SapModel.PropLink.Count

## VB6 Procedure

Function Count(Optional ByVal PropType As eLinkPropType)
As Long

## Parameters

PropType

This optional value is one of the following items in
the eLinkPropType enumeration.

NLPROP\_LINEAR = 1

NLPROP\_DAMPER = 2

NLPROP\_GAP = 3

NLPROP\_HOOK = 4

NLPROP\_PLASTIC\_WEN = 5

NLPROP\_ISOLATOR1 = 6  (Rubber
isolator)

NLPROP\_ISOLATOR2 = 7  (Friction
isolator)

NLPROP\_MULTILINEAR\_ELASTIC
= 8

NLPROP\_MULTILINEAR\_PLASTIC
= 9

NLPROP\_ISOLATOR3 = 10  (T/C
Friction isolator)

If no value is input for PropType, a count is returned
for all link properties in the model regardless of type.

## Remarks

This function returns the total number of defined link
properties in the model. If desired, counts can be returned for all link
properties of a specified type in the model.

## VBA Example



```
Sub CountLinkProps()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

      Dim Count As Long



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
 3, 124, 3, 200)



   'add link property

      ReDim DOF(5)

      ReDim Fixed(5)

      ReDim Ke(5)

      ReDim Ce(5)

      DOF(0) = True

      Ke(0) = 12

      ret = SapModel.PropLink.SetLinear("L1", 
 DOF, Fixed, Ke, Ce, 0, 0)



   'return number of defined links of all types

      Count = SapModel.PropLink.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
