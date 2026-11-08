# SetSpringData

## Syntax

SapObject.SapModel.PropLink.SetSpringData

## VB6 Procedure

Function SetSpringData(ByVal Name As String, ByVal DefinedForThisLength As Double, ByVal DefinedForThisArea As Double) As Long

## Parameters

Name

The name of an existing link property.

DefinedForThisLength

The link property is defined for this length in a line (frame) spring. [L]

DefinedForThisArea

The link property is defined for this area in an area spring. [L2]

## Remarks

This function assigns length and area values to a link property that are used if the link property is specified in line and area spring assignments.

The function returns zero if the values are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetLinkPropSpringData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double



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



   'set link property spring data

      ret = SapModel.PropLink.SetSpringData("L1", 12, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetSpringData](GetSpringData.md)
