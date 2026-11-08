# GetDiscretization

## Syntax

SapObject.SapModel.TendonObj.GetDiscretization

## VB6 Procedure

Function GetDiscretization(ByVal Name As String, ByRef Value As Double) As Long

## Parameters

Name

The name of an existing tendon object.

Value

The maximum discretization length for the tendon. [L]

## Remarks

This function retrieves the maximum discretization length assignment for tendon objects.

The function returns zero if the assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonMaxDiscretizationLength()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Double

      Dim Name As String



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



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'get tendon maximum discretization length

      ret = SapModel.TendonObj.GetDiscretization(Name, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetDiscretization](SetDiscretization.md)
