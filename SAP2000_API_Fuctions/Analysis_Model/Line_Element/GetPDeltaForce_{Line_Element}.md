# GetPDeltaForce

## Syntax

SapObject.SapModel.LineElm.GetPDeltaForce

## VB6 Procedure

Function GetPDeltaForce(ByVal Name As String, ByRef NumberForces As Long, ByRef PDeltaForce() As Double, ByRef Dir() As Long, ByRef CSys() As String) As Long

## Parameters

Name

The name of an existing line element.

NumberForces

The number of P-Delta forces assigned to the line element.

PDeltaForce

This is an array of the P-Delta force values assigned to the line element. [F]

Dir

This is an array that contains 0, 1, 2 or 3, indicating the direction of each P-Delta force assignment.

0 = Frame object local 1-axis direction

1 = Projected X direction in CSys coordinate system

2 = Projected Y direction in CSys coordinate system

3 = Projected Z direction in CSys coordinate system

CSys

This is an array that contains the name of the coordinate system in which each projected P-Delta force is defined. This item is blank when the Dir item is zero, that is, when the P-Delta force is defined in the line element local 1-axis direction.

## Remarks

This function retrieves the P-Delta force assignments to line elements.

The function returns zero if the assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmPDeltaForce()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberForces As Long

      Dim PDeltaForce() As Double

      Dim Dir() As Long

      Dim CSys() As String



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



   'assign P-Delta force to frame object

      ret = SapModel.FrameObj.SetPDeltaForce("ALL", 100, 0, True, , Group)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get P-Delta force for line element

      ret = SapModel.LineElm.GetPDeltaForce("3-1", NumberForces, PDeltaForce, Dir, CSys)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
