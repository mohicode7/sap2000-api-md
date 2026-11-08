# GetMass

## Syntax

SapObject.SapModel.PointObj.GetMass

## VB6 Procedure

Function GetMass(ByVal Name As String, ByRef m() As Double) As Long

## Parameters

Name

The name of an existing point object.

m

This is an array of six mass assignment values.

Value(0) = U1 [M]

Value(1) = U2 [M]

Value(2) = U3 [M]

Value(3) = R1 [ML2]

Value(4) = R2 [ML2]

Value(5) = R3 [ML2]

## Remarks

This function retrieves the point mass assignment values for a point object. The masses are always returned in the point local coordinate system.

The function returns zero if the mass is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i as long

      Dim m() As Double



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



   'assign point mass

      Redim m(5)

      For i = 0 to 5

         m(i) = (i+1) / 10

      Next i

      ret = SapModel.PointObj.SetMass("3", m)



   'get point mass

      Redim m(5)

      ret = SapModel.PointObj.GetMass("3", m)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetMass](SetMass.md)

[DeleteMass](DeleteMass.md)

[SetMassByVolume](SetMassByVolume.md)

[SetMassByWeight](SetMassByWeight.md)
