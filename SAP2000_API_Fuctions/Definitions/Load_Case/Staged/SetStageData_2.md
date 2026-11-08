# SetStageData\_2

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.SetStageData\_2

## VB6 Procedure

Function SetStageData\_2(ByVal Name As String, ByVal
Stage As Long, ByVal NumberOperations As Long, ByRef Operation() As Long,
ByRef ObjectType() As String, ByRef ObjectName() As String, ByRef Age()
As Double, ByRef MyType() As String, ByRef MyName() As String, ByRef SF()
As Double) As Long

## Parameters

Name

The name of an existing static nonlinear staged load
case.

Stage

The stage in the specified load case to which the data
applies. Stages are numbered sequentially starting from 1.

NumberOperations

The number of operations in the specified stage.

Operation

This is an array that includes 1, 2, 3, 4, 5, 6, 7,
or 11, indicating an operation type.

1 = Add structure

2 = Remove structure

3 = Load objects if new

4 = Load objects

5 = Change section properties

6 = Change section property
modifiers

7 = Change releases

11 = Change section properties
and age

ObjectType

This is an array that includes the object type associated
with the specified operation. The object type may be one of the following:

Group

Frame

Cable

Tendon

Area

Solid

Link

Point

The following list shows which object types are applicable
to each operation type:

Operation = 1 (Add structure):  All
object types

Operation = 2 (Remove
structure):  All object types

Operation = 3 (Load objects
if new):  All object types

Operation = 4 (Load objects):  All
object types

Operation = 5 (Change
section properties):  All object types except Point

Operation = 6 (Change
section property modifiers):  Group, Frame, Cable, Area

Operation = 7 (Change releases):  Group,
Frame

Operation = 11 (Change section
properties and age): All object types except Point

ObjectName

This is an array that includes the name of the object
associated with the specified operation. This is the name of a Group,
Frame object, Cable object, Tendon object, Area object, Solid object,
Link object or Point object, depending on the ObjectType item.

Age

This is an array that includes the age of the added
structure, at the time it is added, in days. This item applies only to
operations with Operation = 1.

MyType

This is an array that includes a load type or an object
type, depending on what is specified for the Operation item. This item
applies only to operations with Operation = 3, 4, 5, 6, 7, or 11.

When Operation = 3 or 4, this is an array that includes
Load or Accel, indicating the load type of an added load.

When Operation = 5 or 11, and the ObjectType item is
Group, this is an array that includes Frame, Cable, Tendon, Area, Solid
or Link, indicating the object type for which the section property is
changed.

When Operation = 6 and the ObjectType item is Group,
this is an array that includes Frame, Cable or Area, indicating the object
type for which the section property modifiers are changed.

When Operation = 7 and the ObjectType item is Group,
this is an array that includes Frame, indicating the object type for which
the releases are changed.

When Operation = 5, 6, 7, or 11,  and the ObjectType
item is not Group and not Point, this item is ignored and the type is
picked up from the ObjectType item.

MyName

This is an array that includes a load assignment or
an object name, depending on what is specified for the Operation item.
This item  applies only to operations with Operation = 3, 4, 5, 6,
7 or 11.

When Operation = 3 or 4, this is an array that includes
the name of the load assigned to the operation. If the associated LoadType
item is Load, this item is the name of a defined load pattern. If the
associated LoadType item is Accel, this item is UX, UY, UZ, RX, RY or
RZ, indicating the direction of the load.

When Operation = 5 or 11, this is the name of a Frame,
Cable, Tendon, Area, Solid or Link object, depending on the object type
specified.

When Operation = 6, this is the name of a Frame, Cable
or Area object, depending on the object type specified.

When Operation = 7, this is the name of a Frame object.

SF

This is an array that includes the scale factor for
the load assigned to the operation, if any. [L/s2]
for Accel UX UY and UZ; otherwise unitless

This item applies only to operations with Operation
= 3 or 4.

## Remarks

This function sets the stage data for the specified
stage in the specified load case. All previous stage data for the specified
stage is cleared when this function is called.

The function returns zero if the data is successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearStagedStageData_1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyDuration() As Double

      Dim MyOutput() As Boolean

      Dim MyOutputName() As String

      Dim MyComment() As String

      Dim MyOperation() As Long

      Dim MyObjectType() As String

      Dim MyObjectName() As String

      Dim MyAge() As Double

      Dim MyMyType() As String

      Dim MyMyName() As String

      Dim MySF() As Double



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



   'add static nonlinear staged load case

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("ACASE1")



   'initialize stage definitions

      ReDim MyDuration(1)

      ReDim MyOutput(1)

      ReDim MyOutputName(1)

      ReDim MyComment(1)

      MyDuration(0) = 0

      MyOutput(0) = False

      MyComment(0) = "Build structure"

      MyDuration(1) = 60

      MyOutput(1) = True

      MyOutputName(1) = "HBC2"

      MyComment(1) = "Wait"

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetStageDefinitions_2("ACASE1", 
 2, MyDuration, MyOutput, MyOutputName, MyComment)



   'set stage data

      ReDim MyOperation(1)

      ReDim MyObjectType(1)

      ReDim MyObjectName(1)

      ReDim MyAge(1)

      ReDim MyMyType(1)

      ReDim MyMyName(1)

      ReDim MySF(1)

      MyOperation(0) = 1

      MyObjectType(0) = "Group"

      MyObjectName(0) = "ALL"

      MyAge(0) = 3

      MyOperation(1) = 4

      MyObjectType(1) = "Frame"

      MyObjectName(1) = "8"

      MyMyType(1) = "Load"

      MyMyName(1) = "DEAD"

      MySF(1) = 0.85

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetStageData_2("ACASE1", 
 1, 2, MyOperation, MyObjectType, MyObjectName, MyAge, MyMyType, MyMyName, 
 MySF)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

This function supersedes [SetStageData\_1](../../../Obsolete_Functions/SetStageData_1.md).

## See Also

[GetStageData\_2](GetStageData_2.md)
