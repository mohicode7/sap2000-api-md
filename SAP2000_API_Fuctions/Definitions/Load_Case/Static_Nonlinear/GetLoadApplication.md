# GetLoadApplication

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinear.GetLoadApplication

## VB6 Procedure

Function GetLoadApplication(ByVal Name As String, ByRef LoadControl As Long, ByRef DispType As Long, ByRef Displ As Double, ByRef Monitor As Long, ByRef DOF As Long, ByRef PointName As String, ByRef GDispl As String) As Long

## Parameters

Name

The name of an existing static nonlinear load case.

LoadControl

This is either 1 or 2, indicating the load application control method.

1 = Full load

2 = Displacement control

DispType

This is either 1 or 2, indicating the control displacement type.

1 = Conjugate displacement

2 = Monitored displacement

This item applies only when displacement control is used, that is, LoadControl = 2.

Displ

This item applies only when displacement control is used, that is, LoadControl = 2. The structure is loaded to a monitored displacement of this magnitude. [L] when DOF = 1, 2 or 3 and [rad] when DOF = 4, 5 or 6

Monitor

This is either 1 or 2, indicating the monitored displacement.

1 = Displacement at a specified point object

2 = Generalized displacement

DOF

This is 1, 2, 3, 4, 5 or 6, indicating the degree of freedom for which the displacement at a point object is monitored.

1 = U1

2 = U2

3 = U3

4 = R1

5 = R2

6 = R3

This item applies only when Monitor = 1.

PointName

The name of the point object at which the displacement is monitored. This item applies only when Monitor = 1.

GDispl

The name of the generalized displacement for which the displacement is monitored. This item applies only when Monitor = 2.

## Remarks

This function retrieves the load application control parameters for the specified load case.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearLoadApplication()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SF() As Double

      Dim LoadControl As Long

      Dim DispType As Long

      Dim Displ As Double

      Dim Monitor As Long

      Dim DOF As Long

      Dim PointName As String

      Dim GDispl As String



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



   'add generalized displacement

      ret = SapModel.GDispl.Add("GD1", 1)



   'add point to generalized displacement

      ReDim SF(5)

      SF(0) = 0.5

      ret = SapModel.GDispl.SetPoint("GD1", "2", SF)

      ret = SapModel.GDispl.SetPoint("GD1", "3", SF)



   'add static nonlinear load case

      ret = SapModel.LoadCases.StaticNonlinear.SetCase("LCASE1")



   'set load application parameters

      ret = SapModel.LoadCases.StaticNonlinear.SetLoadApplication("LCASE1", 2, 1, 12, 2, 0, "", "GD1")



   'get load application parameters

      ret = SapModel.LoadCases.StaticNonlinear.GetLoadApplication("LCASE1", LoadControl, DispType, Displ, Monitor, DOF, PointName, GDispl)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SetLoadApplication](GetLoadApplication.md)
