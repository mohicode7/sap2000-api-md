# IsSpringCoupled

## Syntax

SapObject.SapModel.PointObj.IsSpringCoupled

## VB6 Procedure

Function IsSpringCoupled(ByVal Name As String, ByVal IsCoupled As Boolean) As Long

## Parameters

Name

The name of an existing point object.

IsCoupled

This item is True if the spring assigned to the specified point object is coupled, otherwise it is False.

## Remarks

This function indicates if the spring assignments to a point object are coupled, that is, if they have off-diagonal terms in the 6x6 spring matrix for the point object.

The function returns zero if the coupled status is successfully retrieved, otherwise it returns a nonzero value. If no springs exist at the point object,n the function returns a nonzero value.

## VBA Example



```
Sub CheckIsSpringCoupled()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim k() As Double

      Dim IsCoupled As Boolean



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



   'assign spring to a point

      ReDim k(5)

      k(2) = 10

      ret = SapModel.PointObj.SetSpring("3", k)



   'determine if spring is coupled

      ret = SapModel.PointObj.IsSpringCoupled("3", IsCoupled)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BPoint_Object%7D.md)

[GetSpringCoupled](GetSpringCoupled_%7BPoint_Object%7D.md)

[SetSpring](SetSpring.md)

[SetSpringCoupled](SetSpringCoupled.md)

[DeleteSpring](DeleteSpring_%7BPoint_Object%7D.md)
