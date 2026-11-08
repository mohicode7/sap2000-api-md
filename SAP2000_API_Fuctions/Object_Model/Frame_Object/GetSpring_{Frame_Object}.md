# GetSpring

## Syntax

SapObject.SapModel.FrameObj.GetSpring

## VB6 Procedure

Function GetSpring(ByVal Name As String, ByRef NumberSprings As Long, ByRef MyType() As Long, ByRef s() As Double, ByRef SimpleSpringType() As Long, ByRef LinkProp() As String, ByRef SpringLocalOneType() As Long, ByRef Dir() As Long, ByRef Plane23Angle() As Double, ByRef VecX() As Double, ByRef VecY() As Double, ByRef VecZ() As Double, ByRef CSys() As String, ByRef Ang() As Double) As Long

## Parameters

Name

The name of an existing frame object.

NumberSprings

The number of springs assignments made to the specified frame object.

MyType

Each value in this array is either 1 or 2, indicating the spring property type.

1 = Simple spring

2 = Link property

s

Each value in this array is the simple spring stiffness per unit length of the frame object. This item applies only when the corresponding MyType = 1. [F/L2]

SimpleSpringType

Each value in this array is 1, 2 or 3, indicating the simple spring type. This item applies only when the corresponding MyType = 1.

1 = Spring resists tension and compression

2 = Spring resists compression only

3 = Spring resists tension only

LinkProp

Each value in this array is the name of the link property assigned to the spring. This item applies only when the corresponding MyType = 2.

SpringLocalOneType

Each value in this array is 1, 2 or 3, indicating the method used to specify the spring positive local 1-axis orientation.

1 = Parallel to frame object local axis

2 = In the frame object 2-3 plane

3 = User specified direction vector

Dir

Each value in this array is 1, 2, 3, -1, -2 or -3, indicating the frame object local axis that corresponds to the positive local 1-axis of the spring. This item applies only when the corresponding SpringLocalOneType = 1.

Plane23Angle

Each value in this array is the angle in the frame object 2-3 plane measured counter clockwise from the frame positive 2-axis to the spring positive 1-axis. This item applies only when the corresponding SpringLocalOneType = 2. [deg]

VecX

Each value in this array is the X-axis or frame local 1-axis component (depending on the CSys specified) of the user specified direction vector for the spring local 1-axis. The direction vector is in the coordinate system specified by the CSys item. This item applies only when the corresponding SpringLocalOneType = 3.

VecY

Each value in this array is the Y-axis or frame local 2-axis component (depending on the CSys specified) of the user specified direction vector for the spring local 1-axis. The direction vector is in the coordinate system specified by the CSys item. This item applies only when the corresponding SpringLocalOneType = 3.

VecZ

Each value in this array is the X-axis or frame local 3-axis component (depending on the CSys specified) of the user specified direction vector for the spring local 1-axis. The direction vector is in the coordinate system specified by the CSys item. This item applies only when the corresponding SpringLocalOneType = 3.

CSys

Each value in this array is Local (meaning the frame object local coordinate system) or the name of a defined coordinate system. This item is the coordinate system in which the user specified direction vector, Vec, is specified. This item applies only when the corresponding SpringLocalOneType = 3.

Ang

Each value in this array is the angle that the link local 2-axis is rotated from its default orientation. This item applies only when the corresponding MyType = 2. [deg]

## Remarks

This function retrieves the spring assignments to a frame object.

The function returns zero if the assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSprings()

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

      Dim SpringLocalOneType() As Long

      Dim Dir() As Long

      Dim Plane23Angle() As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign springs to frame

      ReDim Vec(2)

      ret = SapModel.FrameObj.SetSpring("ALL", 1, 1, 1, "", 1, 2, 0, Vec, 0, False, "Local", Group)



   'get spring assignments to frames

      ret = SapModel.FrameObj.GetSpring("1", numbersprings, MyType, s, SimpleSpringType, LinkProp, SpringLocalOneType, Dir, Plane23Angle, VecX, VecY, VecZ, CSys, Ang)



   'close Sap2000}

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSpring](SetSpring__%7BFrame_Object%7D.md)

[DeleteSpring](DeleteSpring_%7BFrame_Object%7D.md)
