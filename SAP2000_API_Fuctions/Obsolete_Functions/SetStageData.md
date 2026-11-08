# SetStageData  (Note:  Newer Function Available)

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinear.SetStageData

## VB6 Procedure

Function SetStageData(ByVal Name As String, ByVal Stage As Long, ByVal NumberOperations As Long, ByRef Operation() As Long, ByRef GroupName() As String, ByRef Age() As Long, ByRef LoadType() As String, ByRef LoadName() As String, ByRef SF() As Double) As Long

## Parameters

Name

The name of an existing static nonlinear staged analysis case.

Stage

The stage in the specified load case to which the data applies. Stages are numbered sequentially starting from 1.

NumberOperations

The number of operations in the specified stage.

Operation

This is an array that includes 1, 2, 3 or 4, indicating an operation type.

1 = Add structure

2 = Remove structure

3 = Load added items in group

4 = Load all items in group

GroupName

This is an array that includes the name of the group associated with the specified operation.

Age

This is an array that includes the age of the added structure, at the time it is added, in days. This item applies only to operations with Operation = 1.

LoadType

This is an array that includes either Load or Accel, indicating the load type of an added load. This item applies only to operations with Operation = 3 or 4.

LoadName

This is an array that includes the name of the load assigned to the operation. This item applies only to operations with Operation = 3 or 4.

If the associated LoadType item is Load, this item is the name of a defined load pattern.

If the associated LoadType item is Accel, this item is UX, UY, UZ, RX, RY or RZ, indicating the direction of the load.

SF

This is an array that includes the scale factor for the load assigned to the operation. [L/s2] for Accel UX UY and UZ; otherwise unitless

This item applies only to operations with Operation = 3 or 4.

## Remarks

This function sets the stage data for the specified stage in the specified load case. All previous stage data for the specified stage is cleared when this function is called.

The function returns zero if the data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearStagedStageData()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyDuration() As Long

      Dim MyComment() As String

      Dim MyOperation() As Long

      Dim MyGroupName() As String

      Dim MyAge() As Long

      Dim MyLoadType() As String

      Dim MyLoadName() As String

      Dim MySF() As Double



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add static nonlinear staged load case

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("LCASE1")



   'initialize stage definitions

      ReDim MyDuration(1)

      ReDim MyComment(1)

      MyDuration(0) = 0

      MyComment(0) = "Build structure"

      MyDuration(1) = 60

      MyComment(1) = "Wait"

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetStageDefinitions("LCASE1", 2, MyDuration, MyComment)



   'set stage data

      ReDim MyOperation(1)

      ReDim MyGroupName(1)

      ReDim MyAge(1)

      ReDim MyLoadType(1)

      ReDim MyLoadName(1)

      ReDim MySF(1)

      MyOperation(0) = 1

      MyGroupName(0) = "ALL"

      MyAge(0) = 3

      MyOperation(1) = 4

      MyGroupName(1) = "ALL"

      MyLoadType(1) = "Load"

      MyLoadName(1) = "DEAD"

      MySF(1) = 0.85

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetStageData("LCASE1", 1, 2, MyOperation, MyGroupName, MyAge, MyLoadType, MyLoadName, MySF)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

This function is obsolete and has been superseded by Stag [SetStageData\_1](SetStageData_1.md) as of version 12.00.  This function is maintained for backwards compatibility.

## See Also

[GetStageData](GetStageData.md)
