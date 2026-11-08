# GetMergeNumber

## Syntax

SapObject.SapModel.PointObj.GetMergeNumber

## VB6 Procedure

Function GetMergeNumber(ByVal Name As String, ByRef MergeNumber As Long) As Long

## Parameters

Name

The name of an existing point object.

MergeNumber

The merge number assigned to the specified point object.

## Remarks

This function retrieves the merge number for a point object. By default the merge number for a point is zero. Points with different merge numbers are not automatically merged by the program.

The function returns zero if the merge number is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointMergeNumber()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim m As Long



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



   'set merge number

      ret = SapModel.PointObj.SetMergeNumber("3", 2)



   'get merge number

      ret = SapModel.PointObj.GetMergeNumber("3", m)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetMergeNumber](SetMergeNumber.md)
