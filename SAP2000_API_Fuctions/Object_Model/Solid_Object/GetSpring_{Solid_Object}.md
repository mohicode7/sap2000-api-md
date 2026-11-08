# GetSpring

## Syntax

SapObject.SapModel.SolidObj.GetSpring

## VB6 Procedure

Function GetSpring(ByVal Name As String, ByRef NumberSprings As Long, ByRef MyType() As Long, ByRef s() As Double, ByRef SimpleSpringType() As Long, ByRef LinkProp() As String, ByRef Face() As Long, ByRef SpringLocalOneType() As Long, ByRef Dir() As Long, ByRef Outward() As Boolean, ByRef VecX() As Double, ByRef VecY() As Double, ByRef VecZ() As Double, ByRef CSys() As String, ByRef Ang() As Double) As Long

## Parameters

Name

The name of an existing solid object.

NumberSprings

The number of springs assignments made to the specified solid object.

MyType

Each value in this array is either 1 or 2, indicating the spring property type.

1 = Simple spring

2 = Link property

s

Each value in this array is the simple spring stiffness per unit area of the specified solid object face. This item applies only when the corresponding MyType = 1. [F/L3]

SimpleSpringType

Each value in this array is 1, 2 or 3, indicating the simple spring type. This item applies only when the corresponding MyType = 1.

1 = Spring resists tension and compression

2 = Spring resists compression only

3 = Spring resists tension only

LinkProp

Each value in this array is the name of the link property assigned to the spring. This item applies only when the corresponding MyType = 2.

Face

This is 1, 2, 3, 4, 5 or 6, indicating the solid object face to which the specified spring assignment applies.

SpringLocalOneType

Each value in this array is 1, 2 or 3, indicating the method used to specify the spring positive local 1-axis orientation.

1 = Parallel to solid object local axis

2 = Normal to specified solid object face

3 = User specified direction vector

Dir

Each value in this array is 1, 2, 3, -1, -2 or -3, indicating the solid object local axis that corresponds to the positive local 1-axis of the spring. This item applies only when the corresponding SpringLocalOneType = 1.

Outward

Each value in this array is True if the spring positive local 1 axis is outward from the specified solid object face. This item applies only when SpringLocalOneType = 2.

VecX

Each value in this array is the X-axis or solid object local 1-axis component (depending on the CSys specified) of the user specified direction vector for the spring local 1-axis. The direction vector is in the coordinate system specified by the CSys item. This item applies only when the corresponding SpringLocalOneType = 3.

VecY

Each value in this array is the Y-axis or solid object local 2-axis component (depending on the CSys specified) of the user specified direction vector for the spring local 1-axis. The direction vector is in the coordinate system specified by the CSys item. This item applies only when the corresponding SpringLocalOneType = 3.

VecZ

Each value in this array is the X-axis or solid object local 3-axis component (depending on the CSys specified) of the user specified direction vector for the spring local 1-axis. The direction vector is in the coordinate system specified by the CSys item. This item applies only when the corresponding SpringLocalOneType = 3.

CSys

Each value in this array is Local (meaning the solid object local coordinate system) or the name of a defined coordinate system. This item is the coordinate system in which the user specified direction vector, Vec, is specified. This item applies only when the corresponding SpringLocalOneType = 3.

Ang

Each value in this array is the angle that the link local 2-axis is rotated from its default orientation. This item applies only when the corresponding MyType = 2. [deg]

## Remarks

This function retrieves the spring assignments to a solid object face.

The function returns zero if the assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidObjectSprings()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Vec() As Double

      Dim NumberSprings As Long

      Dim MyType() As Long

      Dim s() As Double

      Dim SimpleSpringType() As Long

      Dim LinkProp() As String

      Dim Face() As Long

      Dim SpringLocalOneType() As Long

      Dim Dir() As Long

      Dim Outward() As Boolean

      Dim VecX() As Double

      Dim VecY() As Double

      Dim VecZ() As Double

      Dim CSys() As String

      Dim Ang() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign springs to solid objects

      ReDim Vec(2)

      ret = SapModel.SolidObj.SetSpring("1", 1, 1, 1, "", 1, 1, 3, True, Vec, 0, False, "Local")



   'get spring assignments to solid objects

      ret = SapModel.SolidObj.GetSpring("1", NumberSprings, MyType, s, SimpleSpringType, LinkProp, Face, SpringLocalOneType, Dir, Outward, VecX, VecY, VecZ, CSys, Ang)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSpring](SetSpring_%7BSolid_Object%7D.md)

[DeleteSpring](DeleteSpring_%7BSolid_Object%7D.md)
