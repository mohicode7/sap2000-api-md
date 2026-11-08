# GetSSCurve

## Syntax

SapObject.SapModel.PropMaterial.GetSSCurve

## VB6 Procedure

Function GetMPIsotropic(ByVal Name As String, ByRef NumberPoints As Long, ByRef PointID() As Long, ByRef Strain() As Double, ByRef Stress() As Double, Optional ByVal SectName As String = "", Optional ByVal RebarArea As Double = 0, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing material property.

NumberPoints

The number of points in the stress-strain curve.

PointID

This is one of the following integers which sets the point ID. The point ID controls the color that will be displayed for hinges in a deformed shape plot.

-5 = -E

-4 = -D

-3 = -C

-2 = -B

0 = None

1 = A

2 = B

3 = C

4 = D

5 = E

Strain

This is an array that includes the strain at each point on the stress strain curve.

Stress

This is an array that includes the stress at each point on the stress strain curve. [F/L2]

SectName

This item applies only if the specified material is concrete with a Mander concrete type.

This is the frame section property for which the Mander stress-strain curve is retrieved.

The section must be round or rectangular.

RebarArea

This item applies only if the specified material is rebar, which does not have a user-defined stress-strain curve and is specified to use Caltrans default controlling strain values, which are bar size dependent.

This is the area of the rebar for which the stress-strain curve is retrieved.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the material stress-strain curve.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example

Sub GetMatPropSSCurve()  
   'dimension variables  
      Dim SapObject as cOAPI  
      Dim SapModel As cSapModel  
      Dim ret As Long  
      Dim NumberPoints As Long  
      Dim PointID() As Long  
      Dim Strain() As Double  
      Dim Stress() As Double  
  
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
  
   'get stress strain curve  
      ret = SapModel.PropMaterial.GetSSCurve("A992Fy50", NumberPoints, PointID, Strain, Stress)  
  
   'close Sap2000  
      SapObject.ApplicationExit False  
      Set SapModel = Nothing  
      Set SapObject = Nothing  
End Sub

## Release Notes

Initial release in version 11.02.

## See Also

[SetSSCurve](SetSSCurve.md)
