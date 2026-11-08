# GetPDelta

## Syntax

SapObject.SapModel.PropLink.GetPDelta

## VB6 Procedure

Function GetPDelta(ByVal Name As String, ByRef Value()
As Double) As Long

## Parameters

Name

The name of an existing link property.

Value

This is an array of P-delta parameters.

Value(0) = M2 P-delta
to I-end of link as moment, M2I

Value(1) = M2 P-delta
to J-end of link as moment, M2J

Value(2) = M3 P-delta
to I-end of link as moment, M3I

Value(3) = M3 P-delta
to J-end of link as moment, M3J

## Remarks

This function retrieves P-delta parameters for a link
property.

The function returns zero if the data is successfully
retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkPropPDelta()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

      Dim MyValue() As Double

      Dim Value() As Double



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



   'add link property

      ReDim DOF(5)

      ReDim Fixed(5)

      ReDim Ke(5)

      ReDim Ce(5)

      DOF(0) = True

      Ke(0) = 12

      ret = SapModel.PropLink.SetLinear("L1", 
 DOF, Fixed, Ke, Ce, 0, 0)



   'set link property P-delta parameters

      ReDim MyValue(3)

      MyValue(0) = 0.6

      MyValue(1) = 0.4

      MyValue(2) = 0.3

      MyValue(3) = 0.2

      ret = SapModel.PropLink.SetPDelta("L1", 
 MyValue)



   'get link property P-delta parameters

      ret = SapModel.PropLink.GetPDelta("L1", 
 Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetPDelta](SetPDelta.md)
